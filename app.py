import io
import os
from pathlib import Path

import streamlit as st

# Propaguj sekrety Streamlit (Streamlit Cloud / .streamlit/secrets.toml) do
# os.environ — extraction_ai i anthropic SDK czytają stamtąd. Bez tego na
# Streamlit Cloud sekrety byłyby dostępne tylko jako st.secrets["..."] i
# aplikacja zgłaszałaby brak klucza.
try:
    for _key in ("ANTHROPIC_API_KEY", "CLAUDE_MODEL"):
        _val = st.secrets.get(_key)
        if _val:
            os.environ.setdefault(_key, str(_val))
except Exception:
    pass  # brak pliku secrets.toml lokalnie — używamy env vars

from main import (
    ASSETS_DIR,
    BASE_DIR,
    FALLBACK_INPUT_DIR,
    INPUT_DIR,
    OUTPUT_DIR,
    TEMP_IMAGES_DIR,
    TEMPLATE_PATH,
    extract_data,
    list_input_pdfs,
    populate_workbook,
    sanitize_filename,
)

st.set_page_config(page_title="Generator Instrukcji BHP", layout="centered")
st.title("Generator Instrukcji BHP")

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

uploaded_files = st.file_uploader(
    "Wgraj pliki PDF z kartami charakterystyki",
    type="pdf",
    accept_multiple_files=True,
)

source_pdfs: list[Path] = []
if uploaded_files:
    for up in uploaded_files:
        pdf_path = INPUT_DIR / up.name
        with open(pdf_path, "wb") as f:
            f.write(up.getbuffer())
        source_pdfs.append(pdf_path)
else:
    source_pdfs = list_input_pdfs()

if not source_pdfs:
    st.warning("Brak plików PDF. Wgraj pliki powyżej lub umieść je w folderze **DANE_WEJSCIOWE** / **DOCS**.")
    st.stop()

st.subheader("Wybierz pliki do wygenerowania")
selected: list[Path] = []
for pdf in source_pdfs:
    rel = str(pdf.relative_to(BASE_DIR)) if BASE_DIR in pdf.parents else pdf.name
    key = f"pdf_{pdf.name}"
    if st.checkbox(rel, value=True, key=key):
        selected.append(pdf)

if not selected:
    st.info("Zaznacz przynajmniej jeden plik.")
    st.stop()

generated_files: list[Path] = []

if not os.environ.get("ANTHROPIC_API_KEY", "").strip():
    st.error(
        "Brak klucza **ANTHROPIC_API_KEY**. Ustaw sekret ANTHROPIC_API_KEY "
        "(w Replit: zakładka **Secrets**) i odśwież stronę — bez klucza "
        "ekstrakcja danych z karty nie zadziała."
    )
    st.stop()

if st.button("Generuj instrukcje"):
    progress = st.progress(0, text="Przetwarzanie...")
    total = len(selected)
    errors = 0

    for i, pdf_path in enumerate(selected):
        try:
            data = extract_data(pdf_path)
            product = data.product_name or pdf_path.stem
            output_name = f"Instrukcja_BHP_{sanitize_filename(product)}.xlsx"
            output_path = OUTPUT_DIR / output_name

            populate_workbook(
                data,
                output_path,
                template_path=TEMPLATE_PATH,
                assets_dir=ASSETS_DIR,
                temp_images_dir=TEMP_IMAGES_DIR,
            )
            generated_files.append(output_path)
        except Exception as exc:  # czytelny komunikat zamiast crashu Streamlit
            errors += 1
            st.error(f"Błąd przetwarzania {pdf_path.name}: {exc}")
        progress.progress((i + 1) / total, text=f"Przetworzono {pdf_path.name}")

    progress.empty()
    if generated_files:
        st.success(f"Wygenerowano {len(generated_files)} plików.")
    if errors:
        st.warning(f"Nie udało się przetworzyć {errors} plików.")

if generated_files:
    st.subheader("Gotowe instrukcje do pobrania")
    cols = st.columns(2)
    for idx, path in enumerate(generated_files):
        with open(path, "rb") as f:
            cols[idx % 2].download_button(
                label=f"Pobierz {path.name}",
                data=f,
                file_name=path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{path.name}",
            )
