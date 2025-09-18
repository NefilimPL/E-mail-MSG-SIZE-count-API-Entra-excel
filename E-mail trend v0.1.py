import os
import sys
import subprocess
import asyncio
import re
import datetime
from collections import defaultdict
from urllib.parse import quote
import msal
import openpyxl
from tqdm import tqdm
import logging

required_modules = ["requests", "msal", "openpyxl", "tqdm", "logging", "aiohttp"]

def install_and_restart():
    try:
        subprocess.check_call([sys.executable.replace('pythonw.exe', 'python.exe'), "-m", "pip", "install"] + required_modules)
        print("Zainstalowano brakujące moduły. Restartowanie skryptu...")
        os.execv(sys.executable.replace('pythonw.exe', 'python.exe'), ['python'] + sys.argv)
    except subprocess.CalledProcessError as e:
        print(f"Błąd podczas instalacji modułów: {e}")
        sys.exit(1)

def check_modules():
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            install_and_restart()

check_modules()

import aiohttp

def safe_int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default

def extract_extended_message_size(message):
    for prop in message.get("singleValueExtendedProperties", []) or []:
        prop_id = (prop.get("id") or "").lower()
        if prop_id == "integer 0x0e08":
            return safe_int(prop.get("value"))
    return 0

LOG_FILENAME = 'email_trend_app_only.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILENAME, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)  # Zmiana na sys.stdout
    ]
)

CLIENT_ID = ""
TENANT_ID = ""
CLIENT_SECRET = ""
SCOPES = ["https://graph.microsoft.com/.default"]

def get_access_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    token_response = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in token_response:
        raise Exception("Nie udało się uzyskać tokena: " + str(token_response))
    return token_response["access_token"]

async def fetch(session, url, headers, semaphore, retries=3, pbar=None):
    async with semaphore:
        while retries > 0:
            try:
                async with session.get(url, headers=headers, timeout=10) as response:
                    if response.status != 200:
                        if pbar:
                            pbar.write(f"Błąd pobierania danych: {response.status} {await response.text()}")
                        retries -= 1
                        if pbar:
                            pbar.write(f"Ponawianie próby pobierania danych... Pozostało prób: {retries}")
                        await asyncio.sleep(5)
                        continue
                    return await response.json()
            except aiohttp.ClientError as e:
                if pbar:
                    pbar.write(f"Błąd połączenia: {e}")
                retries -= 1
                if pbar:
                    pbar.write(f"Ponawianie próby pobierania danych... Pozostało prób: {retries}")
                await asyncio.sleep(5)
            await asyncio.sleep(1)  # Dodaj opóźnienie między żądaniami
    return None

async def fetch_with_error(session, url, headers, semaphore, retries=3, pbar=None):
    async with semaphore:
        attempts = retries
        while attempts > 0:
            try:
                async with session.get(url, headers=headers, timeout=10) as response:
                    if response.status == 200:
                        return await response.json(), None
                    error_body = await response.text()
                    if pbar:
                        pbar.write(f"Błąd pobierania danych: {response.status} {error_body}")
                    if response.status == 400:
                        return None, {"status": response.status, "body": error_body}
                    attempts -= 1
                    if attempts > 0:
                        if pbar:
                            pbar.write(f"Ponawianie próby pobierania danych... Pozostało prób: {attempts}")
                        await asyncio.sleep(5)
                        await asyncio.sleep(1)
                        continue
                    return None, {"status": response.status, "body": error_body}
            except aiohttp.ClientError as e:
                if pbar:
                    pbar.write(f"Błąd połączenia: {e}")
                attempts -= 1
                if attempts > 0:
                    if pbar:
                        pbar.write(f"Ponawianie próby pobierania danych... Pozostało prób: {attempts}")
                    await asyncio.sleep(5)
                    await asyncio.sleep(1)
                    continue
                return None, {"status": None, "body": str(e)}
            await asyncio.sleep(1)
    return None, {"status": None, "body": None}

async def get_child_folders(session, token, mailbox_email, folder, semaphore, path="", pbar=None):
    headers = {"Authorization": f"Bearer {token}"}

    folder_id = folder.get("id")
    folder_name = folder.get("displayName", "")
    total_count = folder.get("totalItemCount", 0)

    if path:
        current_path = f"{path}/{folder_name}"
    else:
        current_path = folder_name

    result = [
        {
            "id": folder_id,
            "path": current_path,
            "displayName": folder_name,
            "totalItemCount": total_count,
        }
    ]

    child_url = f"https://graph.microsoft.com/v1.0/users/{mailbox_email}/mailFolders/{folder_id}/childFolders?$top=100"

    while child_url:
        data = await fetch(session, child_url, headers, semaphore, pbar=pbar)
        if not data:
            break
        children = data.get("value", [])

        for child in children:
            result.extend(await get_child_folders(session, token, mailbox_email, child, semaphore, current_path, pbar))

        child_url = data.get("@odata.nextLink")

    return result

async def get_all_folders(session, token, mailbox_email, semaphore, pbar=None):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox_email}/mailFolders?$top=100"
    all_folders = []

    while url:
        data = await fetch(session, url, headers, semaphore, pbar=pbar)
        if not data:
            raise Exception(f"Błąd pobierania folderów dla {mailbox_email}")
        top_folders = data.get("value", [])

        for f in top_folders:
            all_folders.extend(await get_child_folders(session, token, mailbox_email, f, semaphore, path="", pbar=pbar))

        url = data.get("@odata.nextLink")

    return all_folders

async def get_messages_from_folder(session, token, mailbox_email, folder_id, pbar, semaphore, retries=3):
    headers = {"Authorization": f"Bearer {token}"}
    base_url = (
        "https://graph.microsoft.com/v1.0/users/"
        f"{mailbox_email}/mailFolders/{folder_id}/messages"
    )
    attachments_expand = "attachments($select=size)"
    extended_filter = quote("id eq 'Integer 0x0E08'", safe="")

    def build_url():
        select_parts = [
            "subject",
            "receivedDateTime",
            "hasAttachments",
            "id",
            "from",
            "singleValueExtendedProperties",
        ]
        expand_parts = [
            attachments_expand,
            f"singleValueExtendedProperties($filter={extended_filter})",
        ]
        select_clause = ",".join(select_parts)
        expand_clause = ",".join(expand_parts)
        return f"{base_url}?$select={select_clause}&$expand={expand_clause}&$top=100"

    url = build_url()
    size_warning_logged = False
    messages = []

    while url:
        data, error = await fetch_with_error(session, url, headers, semaphore, retries, pbar)
        if not data:
            error_body = ((error or {}).get("body") or "").lower()
            if (
                (error or {}).get("status") == 400
                and "property named 'size'" in error_body
                and not size_warning_logged
            ):
                logging.warning(
                    "API Graph odrzuciło właściwość size dla skrzynki %s. "
                    "Wszystkie rozmiary wiadomości będą pobierane z rozszerzonej właściwości PR_MESSAGE_SIZE.",
                    mailbox_email,
                )
                size_warning_logged = True
            break

        page_messages = data.get("value", [])
        for msg in page_messages:
            attachments = msg.get("attachments", []) or []
            attachment_size = sum(safe_int(att.get("size")) for att in attachments)
            msg["attachment_size"] = attachment_size

            total_size = extract_extended_message_size(msg)
            if total_size <= 0:
                total_size = safe_int(msg.get("size"))
            if total_size <= 0:
                total_size = attachment_size

            body_size = max(total_size - attachment_size, 0)

            msg["total_size"] = total_size
            msg["body_size"] = body_size
            msg["size"] = body_size

            msg.pop("attachments", None)
            msg.pop("singleValueExtendedProperties", None)

        messages.extend(page_messages)
        pbar.update(len(page_messages))
        url = data.get("@odata.nextLink")

    return messages

def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r'[\\/:\?\*\[\]]+', '_', name)
    return cleaned[:31] or "Folder"

def build_monthly_summary(mailbox_data):
    summary = defaultdict(
        lambda: {
            "message_count": 0,
            "body_size": 0,
            "attachment_size": 0,
            "total_size": 0,
        }
    )
    for folder_path, messages in mailbox_data.items():
        for msg in messages:
            body_bytes = safe_int(msg.get("body_size", msg.get("size", 0)))
            attachment_bytes = safe_int(msg.get("attachment_size", 0))
            total_bytes = safe_int(
                msg.get("total_size", body_bytes + attachment_bytes)
            )
            received_dt = msg.get("receivedDateTime")
            month_key = "Nieznany"
            if received_dt:
                try:
                    dt_str = received_dt.replace("Z", "")
                    dt_obj = datetime.datetime.fromisoformat(dt_str)
                    month_key = dt_obj.strftime("%Y-%m")
                except ValueError:
                    month_key = received_dt[:7]

            summary[(folder_path, month_key)]["message_count"] += 1
            summary[(folder_path, month_key)]["body_size"] += body_bytes
            summary[(folder_path, month_key)]["attachment_size"] += attachment_bytes
            summary[(folder_path, month_key)]["total_size"] += total_bytes

    return summary

def export_to_excel(data, mailbox_email):
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        default_sheet = wb["Sheet"]
        wb.remove(default_sheet)

    for folder_path, messages in data.items():
        sheet_name = sanitize_sheet_name(folder_path)

        ws = wb.create_sheet(title=sheet_name)
        ws.append([
            "Subject",
            "Sender",
            "Message Size (bytes)",
            "Message Size (KB)",
            "Message Size (MB)",
            "Attachment Size (bytes)",
            "Attachment Size (KB)",
            "Attachment Size (MB)",
            "Total Size (bytes)",
            "Total Size (KB)",
            "Total Size (MB)",
            "Has Attachments",
            "Received Date",
            "Received Time",
            "Month"
        ])

        for msg in messages:
            subject = msg.get("subject")
            sender = msg.get("from", {}).get("emailAddress", {}).get("address", "")

            body_bytes = safe_int(msg.get("body_size", msg.get("size", 0)))
            body_kb = round(body_bytes / 1024, 2)
            body_mb = round(body_bytes / (1024 * 1024), 2)

            attach_bytes = safe_int(msg.get("attachment_size", 0))
            attach_kb = round(attach_bytes / 1024, 2)
            attach_mb = round(attach_bytes / (1024 * 1024), 2)

            total_bytes = safe_int(msg.get("total_size", body_bytes + attach_bytes))
            total_kb = round(total_bytes / 1024, 2)
            total_mb = round(total_bytes / (1024 * 1024), 2)

            has_attachments = "Yes" if attach_bytes > 0 else "No"

            received_dt = msg.get("receivedDateTime")
            month_label = ""
            if received_dt:
                try:
                    dt_str = received_dt.replace("Z", "")
                    dt_obj = datetime.datetime.fromisoformat(dt_str)
                    received_date = dt_obj.date().strftime("%Y-%m-%d")
                    received_time = dt_obj.time().strftime("%H:%M:%S")
                    month_label = dt_obj.strftime("%Y-%m")
                except ValueError:
                    received_date = received_dt
                    received_time = ""
                    month_label = received_dt[:7]
            else:
                received_date = ""
                received_time = ""
                month_label = "Nieznany"

            ws.append([
                subject,
                sender,
                body_bytes,
                body_kb,
                body_mb,
                attach_bytes,
                attach_kb,
                attach_mb,
                total_bytes,
                total_kb,
                total_mb,
                has_attachments,
                received_date,
                received_time,
                month_label
            ])

    summary_sheet = wb.create_sheet(title="Podsumowanie")
    summary_sheet.append([
        "Mailbox",
        "Folder",
        "Month",
        "Message Count",
        "Total Size (KB)",
        "Message Size (KB)",
        "Attachment Size (KB)",
        "Total Size (MB)",
        "Message Size (MB)",
        "Attachment Size (MB)"
    ])

    summary_data = build_monthly_summary(data)
    for (folder_path, month_key), values in sorted(summary_data.items(), key=lambda x: (x[0][0], x[0][1])):
        total_size_bytes = safe_int(values.get("total_size", 0))
        body_size_bytes = safe_int(values.get("body_size", 0))
        attachment_size_bytes = safe_int(values.get("attachment_size", 0))
        summary_sheet.append([
            mailbox_email,
            folder_path,
            month_key,
            values["message_count"],
            round(total_size_bytes / 1024, 2),
            round(body_size_bytes / 1024, 2),
            round(attachment_size_bytes / 1024, 2),
            round(total_size_bytes / (1024 * 1024), 2),
            round(body_size_bytes / (1024 * 1024), 2),
            round(attachment_size_bytes / (1024 * 1024), 2)
        ])

    safe_mailbox = mailbox_email.replace("@", "_at_").replace(".", "_")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{safe_mailbox}_{timestamp}.xlsx"
    wb.save(filename)
    logging.info(f"Dane zapisano do pliku: {filename}")

async def process_mailbox(session, mailbox, token, semaphore):
    try:
        logging.info(f"Przetwarzanie skrzynki: {mailbox}")
        with tqdm(total=1, desc=f"Przetwarzanie {mailbox}", unit="msg", position=0, leave=True) as pbar:
            folders = await get_all_folders(session, token, mailbox, semaphore, pbar)
            total_msgs = sum(f.get("totalItemCount", 0) for f in folders)
            pbar.total = total_msgs

            mailbox_data = {}

            tasks = [get_messages_from_folder(session, token, mailbox, f["id"], pbar, semaphore) for f in folders]
            results = await asyncio.gather(*tasks)
            for f, messages in zip(folders, results):
                mailbox_data[f["path"]] = messages

            export_to_excel(mailbox_data, mailbox)
    except Exception as e:
        logging.error(f"Błąd przetwarzania skrzynki {mailbox}: {str(e)}")

async def main():
    logging.info("Rozpoczynam pobieranie danych (app-only)...")
    token = get_access_token()
    logging.info("Token dostępu uzyskany pomyślnie.")

    mailboxes_input = input("Podaj adresy skrzynek oddzielone przecinkiem: ").strip()
    mailbox_list = [m.strip() for m in mailboxes_input.split(",") if m.strip()]

    semaphore = asyncio.Semaphore(7)  # Zwiększ liczbę jednoczesnych żądań

    async with aiohttp.ClientSession() as session:
        tasks = [process_mailbox(session, mailbox, token, semaphore) for mailbox in mailbox_list]
        await asyncio.gather(*tasks)

    logging.info("Przetwarzanie zakończone.")

if __name__ == "__main__":
    asyncio.run(main())
