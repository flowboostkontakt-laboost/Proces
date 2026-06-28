#!/usr/bin/env python3
"""Tworzy dane nowego użytkownika do logowania w aplikacji.

Pyta o login, nazwę, e-mail i hasło, generuje HASH bcrypt (hasła nigdzie nie
zapisujemy w postaci jawnej!) i wypisuje gotowy fragment TOML do wklejenia
w sekretach aplikacji (Streamlit Cloud → Settings → Secrets) albo do pliku
`.streamlit/secrets.toml`.

Uruchomienie:
    python tworz_uzytkownika.py
"""
import getpass
import sys

try:
    import bcrypt
except ModuleNotFoundError:
    sys.exit(
        "Brak biblioteki bcrypt. Zainstaluj: pip install -r requirements.txt"
    )


def main() -> None:
    print("== Nowy użytkownik aplikacji ==")
    username = input("Login (bez spacji, np. jkowalski): ").strip()
    if not username:
        sys.exit("Login nie może być pusty.")

    name = input("Imię i nazwisko / nazwa wyświetlana: ").strip() or username
    email = input("E-mail (opcjonalnie, Enter aby pominąć): ").strip()

    pw1 = getpass.getpass("Hasło: ")
    if not pw1:
        sys.exit("Hasło nie może być puste.")
    pw2 = getpass.getpass("Powtórz hasło: ")
    if pw1 != pw2:
        sys.exit("Hasła się różnią — uruchom skrypt ponownie.")

    hashed = bcrypt.hashpw(pw1.encode(), bcrypt.gensalt()).decode()

    print("\n--- Wklej poniższy fragment do sekcji [auth] w sekretach ---\n")
    print(f"[auth.credentials.usernames.{username}]")
    print(f'name = "{name}"')
    print(f'email = "{email}"')
    print(f'password = "{hashed}"')
    print()


if __name__ == "__main__":
    main()
