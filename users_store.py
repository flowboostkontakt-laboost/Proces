"""Trwały magazyn użytkowników w Supabase (PostgreSQL).

Konta dodawane z panelu administratora zapisują się tutaj i przeżywają
restart aplikacji (w przeciwieństwie do plików na Streamlit Cloud).
Tabela `app_users` tworzona jest automatycznie przy pierwszym połączeniu.

Konfiguracja w sekretach:

    [supabase]
    uri = "postgresql://postgres.<ref>:HASLO@aws-0-<region>.pooler.supabase.com:5432/postgres"

Użyj connection stringa z trybu **Session pooler** (Supabase → Connect),
bo jest dostępny po IPv4 — tak łączy się Streamlit Cloud.
"""
from __future__ import annotations

import bcrypt
import streamlit as st
from sqlalchemy import create_engine, text

TABLE = "app_users"
MIN_PASSWORD_LEN = 6


def _get_uri() -> str | None:
    if "supabase" in st.secrets:
        return st.secrets["supabase"].get("uri") or None
    return st.secrets.get("DATABASE_URL") or None


def _normalize(uri: str) -> str:
    # SQLAlchemy wymaga schematu postgresql:// (Supabase czasem podaje postgres://)
    if uri.startswith("postgres://"):
        uri = "postgresql://" + uri[len("postgres://") :]
    return uri


def is_configured() -> bool:
    return bool(_get_uri())


@st.cache_resource(show_spinner=False)
def _engine(uri: str):
    eng = create_engine(_normalize(uri), pool_pre_ping=True)
    with eng.begin() as conn:
        conn.execute(
            text(
                f"CREATE TABLE IF NOT EXISTS {TABLE} ("
                "username TEXT PRIMARY KEY, "
                "name TEXT, "
                "email TEXT, "
                "password TEXT NOT NULL, "
                "is_admin BOOLEAN NOT NULL DEFAULT FALSE)"
            )
        )
    return eng


def _eng():
    uri = _get_uri()
    if not uri:
        raise RuntimeError("Brak konfiguracji Supabase (sekcja [supabase] uri w sekretach).")
    return _engine(uri)


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()


def _row_to_user(r) -> dict:
    return {
        "username": r["username"],
        "name": r["name"] or r["username"],
        "email": r["email"] or "",
        "is_admin": bool(r["is_admin"]),
        "password": r["password"],
    }


_COLS = "username, name, email, password, is_admin"


def list_users() -> list[dict]:
    with _eng().connect() as conn:
        rows = conn.execute(
            text(f"SELECT {_COLS} FROM {TABLE} ORDER BY username")
        ).mappings().all()
    return [_row_to_user(r) for r in rows]


def get_user(username: str) -> dict | None:
    with _eng().connect() as conn:
        r = conn.execute(
            text(f"SELECT {_COLS} FROM {TABLE} WHERE username = :u"),
            {"u": username},
        ).mappings().first()
    return _row_to_user(r) if r else None


def add_user(username: str, name: str, email: str, plain: str, is_admin: bool = False) -> None:
    username = (username or "").strip()
    if not username or " " in username:
        raise ValueError("Login nie może być pusty ani zawierać spacji.")
    if not plain or len(plain) < MIN_PASSWORD_LEN:
        raise ValueError(f"Hasło musi mieć co najmniej {MIN_PASSWORD_LEN} znaków.")
    with _eng().begin() as conn:
        exists = conn.execute(
            text(f"SELECT 1 FROM {TABLE} WHERE username = :u"), {"u": username}
        ).first()
        if exists:
            raise ValueError(f"Użytkownik '{username}' już istnieje.")
        conn.execute(
            text(
                f"INSERT INTO {TABLE} ({_COLS}) "
                "VALUES (:u, :n, :e, :p, :a)"
            ),
            {
                "u": username,
                "n": (name or "").strip() or username,
                "e": (email or "").strip(),
                "p": hash_password(plain),
                "a": bool(is_admin),
            },
        )


def set_password(username: str, plain: str) -> None:
    if not plain or len(plain) < MIN_PASSWORD_LEN:
        raise ValueError(f"Hasło musi mieć co najmniej {MIN_PASSWORD_LEN} znaków.")
    with _eng().begin() as conn:
        res = conn.execute(
            text(f"UPDATE {TABLE} SET password = :p WHERE username = :u"),
            {"p": hash_password(plain), "u": username},
        )
        if res.rowcount == 0:
            raise ValueError(f"Nie ma użytkownika '{username}'.")


def set_admin(username: str, is_admin: bool) -> None:
    with _eng().begin() as conn:
        conn.execute(
            text(f"UPDATE {TABLE} SET is_admin = :a WHERE username = :u"),
            {"a": bool(is_admin), "u": username},
        )


def delete_user(username: str) -> None:
    with _eng().begin() as conn:
        res = conn.execute(
            text(f"DELETE FROM {TABLE} WHERE username = :u"), {"u": username}
        )
        if res.rowcount == 0:
            raise ValueError(f"Nie ma użytkownika '{username}'.")
