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
st.markdown("Wybierz pliki PDF do przetworzenia, a następnie kliknij **Generuj instrukcje**.")

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

pdfs = list_input_pdfs()

if not pdfs:
    st.warning("Brak plików PDF w folderach **DANE_WEJSCIOWE** ani **DOCS**.")
    st.stop()

st.subheader("Wybierz pliki do wygenerowania")
selected: list[Path] = []
for pdf in pdfs:
    rel = str(pdf.relative_to(BASE_DIR))
    if st.checkbox(rel, value=True, key=rel):
        selected.append(pdf)

if not selected:
    st.info("Zaznacz przynajmniej jeden plik.")
    st.stop()

generated_files: list[Path] = []

if st.button("Generuj instrukcje"):
    progress = st.progress(0, text="Przetwarzanie...")
    total = len(selected)

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
    st.success(f"Wygenerowano {len(generated_files)} plików w folderze **GOTOWE_INSTRUKCJE**.")

if generated_files:
    st.subheader("Gotowe instrukcje do pobrania")
    for path in generated_files:
        with open(path, "rb") as f:
            st.download_button(
                label=f"Pobierz {path.name}",
                data=f,
                file_name=path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
