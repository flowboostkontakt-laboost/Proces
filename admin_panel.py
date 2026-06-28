"""Panel administratora — dodawanie / edycja / usuwanie użytkowników.

Renderowany tylko dla zalogowanych administratorów (patrz auth.is_admin).
Konta zapisywane są w Supabase / PostgreSQL (users_store).
"""
from __future__ import annotations

import streamlit as st

import users_store


def render(current_username: str) -> None:
    with st.expander("👤 Zarządzanie użytkownikami (administrator)", expanded=False):
        if not users_store.is_configured():
            st.warning(
                "Trwały magazyn użytkowników (Supabase) nie jest skonfigurowany. "
                "Dodaj sekcję [supabase] z polem `uri` w sekretach aplikacji, aby "
                "dodawać konta z poziomu aplikacji."
            )
            return

        try:
            users = users_store.list_users()
        except Exception as exc:  # połączenie / connection string
            st.error(
                "Nie udało się połączyć z bazą użytkowników. Sprawdź [supabase].uri "
                "w sekretach (connection string z trybu Session pooler: "
                "Supabase → Connect)."
            )
            st.caption(f"Szczegóły techniczne: {exc}")
            return

        # --- Dodawanie użytkownika ---
        st.markdown("**➕ Dodaj użytkownika**")
        with st.form("add_user_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            new_username = c1.text_input("Login *", placeholder="np. jkowalski")
            new_name = c2.text_input("Imię i nazwisko")
            c3, c4 = st.columns(2)
            new_email = c3.text_input("E-mail")
            new_pwd = c4.text_input("Hasło *", type="password")
            new_admin = st.checkbox("Uprawnienia administratora")
            if st.form_submit_button("Dodaj użytkownika"):
                try:
                    users_store.add_user(new_username, new_name, new_email, new_pwd, new_admin)
                    st.success(f"Dodano użytkownika '{new_username.strip()}'.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        st.divider()

        # --- Lista istniejących użytkowników ---
        st.markdown(f"**Użytkownicy w bazie ({len(users)})**")
        if not users:
            st.caption("Brak kont w bazie. Konto awaryjne 'admin' znajduje się w sekretach.")

        for u in users:
            uname = u["username"]
            title = f"{uname} — {u['name']}" + ("  🔑 admin" if u["is_admin"] else "")
            with st.expander(title):
                st.caption(f"E-mail: {u['email'] or '—'}")

                with st.form(f"reset_{uname}", clear_on_submit=True):
                    new_pw = st.text_input("Nowe hasło", type="password", key=f"np_{uname}")
                    if st.form_submit_button("Zmień hasło"):
                        try:
                            users_store.set_password(uname, new_pw)
                            st.success("Hasło zmienione.")
                        except Exception as exc:
                            st.error(str(exc))

                col_a, col_b = st.columns(2)
                toggle_label = "Odbierz uprawnienia admina" if u["is_admin"] else "Nadaj uprawnienia admina"
                if col_a.button(toggle_label, key=f"adm_{uname}"):
                    try:
                        users_store.set_admin(uname, not u["is_admin"])
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))

                if uname == current_username:
                    col_b.caption("To Twoje konto — nie możesz go usunąć.")
                elif col_b.button("🗑️ Usuń użytkownika", key=f"del_{uname}"):
                    try:
                        users_store.delete_user(uname)
                        st.warning(f"Usunięto użytkownika '{uname}'.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
