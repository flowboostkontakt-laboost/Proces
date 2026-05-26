import io
from pathlib import Path

import streamlit as st

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

ADMIN_CONTACT_MESSAGE = (
    "Wystąpił problem podczas generowania instrukcji. "
    "Skontaktuj się z administratorem."
)

if st.button("Generuj instrukcje"):
    progress = st.progress(0, text="Przetwarzanie...")
    total = len(selected)
    generated_files: list[Path] = []

    try:
        for i, pdf_path in enumerate(selected):
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
            progress.progress((i + 1) / total, text=f"Przetworzono {pdf_path.name}")

        progress.empty()
        st.session_state["generated_files"] = [str(p) for p in generated_files]
        st.success(f"Wygenerowano {len(generated_files)} plików.")
    except Exception:
        progress.empty()
        st.session_state["generated_files"] = [str(p) for p in generated_files]
        st.error(ADMIN_CONTACT_MESSAGE)

generated_files = [Path(p) for p in st.session_state.get("generated_files", [])]
generated_files = [p for p in generated_files if p.exists()]

if generated_files:
    st.subheader("Gotowe instrukcje do pobrania")
    cols = st.columns(2)
    for idx, path in enumerate(generated_files):
        with open(path, "rb") as f:
            cols[idx % 2].download_button(
                label=f"Pobierz {path.name}",
                data=f.read(),
                file_name=path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{path.name}",
            )
