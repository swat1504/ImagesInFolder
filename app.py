import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO

st.set_page_config(page_title="Image → Word Converter")

st.title("📄 Image Folder → Word File Generator")

uploaded_files = st.file_uploader(
    "Upload Images",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:

    uploaded_files = sorted(uploaded_files, key=lambda x: x.name)

    if st.button("Generate Word File"):

        doc = Document()

        for i in range(0, len(uploaded_files), 4):

            table = doc.add_table(rows=2, cols=2)

            chunk = uploaded_files[i:i+4]
            idx = 0

            for r in range(2):
                for c in range(2):
                    if idx < len(chunk):
                        cell = table.rows[r].cells[c]
                        cell.paragraphs[0].add_run().add_picture(
                            chunk[idx],
                            width=Inches(3)
                        )
                        idx += 1

            if i + 4 < len(uploaded_files):
                doc.add_page_break()

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Done!")

        st.download_button(
            label="⬇ Download Word File",
            data=buffer,
            file_name="images.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
