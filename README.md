# E-mail-MSG-SIZE-count-API-Entra-excel
https://entra.microsoft.com/#home

Uprawnienia w Entrze dla klucza API:
##### Mail.Read  Delegowane  Odczytuj pocztę użytkownika
##### Mail.Read  Aplikacja  Read mail in all mailboxes
##### User.Read  Delegowane  Loguj się i odczytuj profil użytkownika
<img width="1029" height="317" alt="msedge_o87vDLrEiz" src="https://github.com/user-attachments/assets/e0e30511-fe3a-46aa-9c9e-c89e159e1959" />

## Konfiguracja

Skrypt automatycznie sprawdza obecność pliku `email_trend_config.json` w tym samym katalogu, w którym znajduje się skrypt Python. Jeśli plik nie istnieje, zostanie wygenerowany szablon z wartościami domyślnymi. W takiej sytuacji należy:

1. Uruchomić skrypt (`python "E-mail trend v0.1.py"`).
2. Po pierwszym uruchomieniu pojawi się plik `email_trend_config.json`.
3. Uzupełnić pola `client_id`, `tenant_id` oraz `client_secret` danymi z aplikacji w Entra ID.
4. Opcjonalnie dopasować pozostałe ustawienia (zakresy uprawnień, poziom logowania, limity czasowe, liczbę równoległych zapytań i rozmiar paczek folderów).
5. Zapisać zmiany i ponownie uruchomić skrypt.

### Przykładowa struktura pliku `email_trend_config.json`

```json
{
  "client_id": "00000000-0000-0000-0000-000000000000",
  "tenant_id": "00000000-0000-0000-0000-000000000000",
  "client_secret": "super_tajne_haslo",
  "scopes": [
    "https://graph.microsoft.com/.default"
  ],
  "log_filename": "email_trend_app_only.log",
  "log_level": "INFO",
  "fetch_timeout_seconds": 30,
  "retry_delay_seconds": 5,
  "throttle_delay_seconds": 1,
  "semaphore_limit": 7,
  "max_folder_batch_size": 3
}
```

## Jak działa skrypt

1. **Kontrola środowiska** – przy pierwszym uruchomieniu skrypt sprawdza, czy wymagane moduły (`requests`, `msal`, `openpyxl`, `tqdm`, `aiohttp`) są dostępne. Brakujące biblioteki są instalowane automatycznie, a skrypt wznawia działanie po zakończeniu instalacji.
2. **Ładowanie konfiguracji** – plik `email_trend_config.json` jest wczytywany i walidowany. Brakujące klucze są dopisywane z wartościami domyślnymi, a nieprawidłowe wartości (np. ujemne limity czasowe) są zastępowane bezpiecznymi ustawieniami.
3. **Uwierzytelnianie** – na podstawie `client_id`, `tenant_id`, `client_secret` i listy `scopes` tworzony jest klient MSAL, który pobiera token dostępu aplikacji (tryb app-only) do Microsoft Graph.
4. **Pobieranie skrzynek** – po podaniu adresów e-mail skrypt równolegle przetwarza każdą skrzynkę. Dla każdej skrzynki rekurencyjnie pobiera strukturę folderów, korzystając z ograniczeń `semaphore_limit` oraz opóźnień `throttle_delay_seconds`, aby nie przeciążać API.
5. **Pobieranie wiadomości** – z każdego folderu pobierane są wiadomości wraz z nagłówkami, rozmiarem ciała i załączników. Skrypt potrafi oszacować rozmiar wiadomości nawet wtedy, gdy Graph nie zwraca wszystkich danych, np. na podstawie nagłówków i podglądu treści.
6. **Obsługa błędów** – operacje sieciowe mają wbudowane ponawianie (`retry_delay_seconds`) i limit czasu (`fetch_timeout_seconds`). Każda nieudana próba jest logowana, a skrócone komunikaty błędów pozwalają szybko znaleźć przyczynę problemu.
7. **Eksport do Excela** – po zebraniu wszystkich wiadomości dane zapisywane są do pliku `.xlsx`. Powstaje osobna karta dla każdego folderu (z listą wiadomości i rozmiarami) oraz karta `Podsumowanie`, która agreguje liczbę wiadomości i łączny rozmiar miesięcznie dla każdego folderu.
8. **Informacje pomocnicze** – pasek postępu (`tqdm`) pokazuje liczbę przetworzonych wiadomości, a logi zapisywane są zarówno do pliku jak i na standardowe wyjście, co ułatwia nadzór nad działaniem narzędzia.


### Logowanie

* Logi są zapisywane do pliku wskazanego w `log_filename` (domyślnie `email_trend_app_only.log` w katalogu skryptu) oraz wypisywane na standardowe wyjście.
* Poziom logowania można zmienić w polu `log_level` (np. `DEBUG`, `INFO`, `WARNING`).
* Błędy związane z pobieraniem danych są skracane do czytelnej formy, aby logi zawierały jak najwięcej przydatnych informacji, ale jednocześnie pozostawały zwięzłe.
