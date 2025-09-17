import os
import sys
import subprocess
import asyncio
import re
import datetime
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

async def get_messages_from_folder(session, token, mailbox_email, folder_id, pbar, semaphore, url=None, retries=3):
    headers = {"Authorization": f"Bearer {token}"}
    if url is None:
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_email}/mailFolders/{folder_id}/messages?$select=subject,receivedDateTime,body,hasAttachments,id,internetMessageHeaders,from"

    messages = []

    while url:
        data = await fetch(session, url, headers, semaphore, retries, pbar)
        if not data:
            break
        page_messages = data.get("value", [])
        for msg in page_messages:
            msg_size, attachment_size = await estimate_message_size(session, token, mailbox_email, msg, semaphore, pbar)
            msg["size"] = msg_size
            msg["attachment_size"] = attachment_size
        messages.extend(page_messages)
        pbar.update(len(page_messages))
        url = data.get("@odata.nextLink")

    return messages

async def estimate_message_size(session, token, mailbox_email, message, semaphore, pbar=None):
    body_content = message.get("body", {}).get("content", "")
    body_size = len(body_content.encode('utf-8')) if body_content else 1024
    attachment_size = await get_attachments_size(session, token, mailbox_email, message.get("id"), semaphore, pbar) if message.get("hasAttachments") else 0
    header_size = sum(len(h.get("name", "").encode('utf-8')) + len(h.get("value", "").encode('utf-8')) for h in message.get("internetMessageHeaders", []))
    total_size = body_size + attachment_size + header_size
    return total_size, attachment_size

async def get_attachments_size(session, token, mailbox_email, message_id, semaphore, pbar=None):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox_email}/messages/{message_id}/attachments?$select=size"
    data = await fetch(session, url, headers, semaphore, pbar=pbar)
    if not data:
        return 0
    return sum(att.get("size", 0) for att in data.get("value", []))

def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r'[\\/:\?\*\[\]]+', '_', name)
    return cleaned[:31] or "Folder"

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
            "Size (bytes)",
            "Size (KB)",
            "Size (MB)",
            "Attachment Size (bytes)",
            "Attachment Size (KB)",
            "Attachment Size (MB)",
            "Has Attachments",
            "Received Date",
            "Received Time"
        ])

        for msg in messages:
            subject = msg.get("subject")
            sender = msg.get("from", {}).get("emailAddress", {}).get("address", "")

            size_bytes = msg.get("size", 0)
            size_kb = round(size_bytes / 1024, 2)
            size_mb = round(size_bytes / (1024 * 1024), 2)

            attach_bytes = msg.get("attachment_size", 0)
            attach_kb = round(attach_bytes / 1024, 2)
            attach_mb = round(attach_bytes / (1024 * 1024), 2)

            has_attachments = "Yes" if attach_bytes > 0 else "No"

            received_dt = msg.get("receivedDateTime")
            if received_dt:
                try:
                    dt_str = received_dt.replace("Z", "")
                    dt_obj = datetime.datetime.fromisoformat(dt_str)
                    received_date = dt_obj.date().strftime("%Y-%m-%d")
                    received_time = dt_obj.time().strftime("%H:%M:%S")
                except ValueError:
                    received_date = received_dt
                    received_time = ""
            else:
                received_date = ""
                received_time = ""

            ws.append([
                subject,
                sender,
                size_bytes,
                size_kb,
                size_mb,
                attach_bytes,
                attach_kb,
                attach_mb,
                has_attachments,
                received_date,
                received_time
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
