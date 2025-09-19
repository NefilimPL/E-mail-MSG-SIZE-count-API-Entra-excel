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

### Ograniczanie throttlingu

* `semaphore_limit` określa maksymalną liczbę równoległych żądań HTTP, jakie mogą być wykonywane jednocześnie.
* `max_folder_batch_size` ogranicza liczbę folderów pobieranych w jednej paczce, co zmniejsza krótkotrwałe skoki obciążenia.
* Skrypt respektuje odpowiedzi 429 (`Retry-After`), wprowadza wykładniczy backoff i współdzieloną kolejkę żądań, dzięki czemu kolejne zapytania są automatycznie spowalniane po sygnale o limitach.

### Logowanie

* Logi są zapisywane do pliku wskazanego w `log_filename` (domyślnie `email_trend_app_only.log` w katalogu skryptu) oraz wypisywane na standardowe wyjście.
* Poziom logowania można zmienić w polu `log_level` (np. `DEBUG`, `INFO`, `WARNING`).
* Błędy związane z pobieraniem danych są skracane do czytelnej formy, aby logi zawierały jak najwięcej przydatnych informacji, ale jednocześnie pozostawały zwięzłe.
