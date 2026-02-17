import streamlit as st
import os
import tempfile
from docx import Document
from docx.shared import Inches

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, Image, Spacer, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

PAGE_WIDTH, PAGE_HEIGHT = A4

VISIT_TYPES = [
    "RESIDENCE",
    "EMPLOYMENT / BUSINESS",
    "RESI CUM OFFICE",
    "Others"
]

def safe_name(text):
    return text.replace("/", "_").replace(" ", "_")

# ---------------- PDF ---------------- #

def generate_pdf(visit_images, output_path):

    LEFT = 40
    RIGHT = 40
    TOP = 40
    BOTTOM = 80

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=LEFT,
        rightMargin=RIGHT,
        topMargin=TOP,
        bottomMargin=BOTTOM
    )

    styles = getSampleStyleSheet()
    elements = []

    usable_width = PAGE_WIDTH - LEFT - RIGHT
    usable_height = PAGE_HEIGHT - TOP - BOTTOM

    img_w = usable_width / 2 - 12
    img_h = usable_height / 2 - 40

    def resize(path):
        img = Image(path)
        scale = min(img_w/img.drawWidth, img_h/img.drawHeight)
        img.drawWidth *= scale
        img.drawHeight *= scale
        return img

    first_section = True

    for visit, images in visit_images.items():

        if not images:
            continue

        if not first_section:
            elements.append(PageBreak())

        first_section = False

        title = visit if visit == "Others" else f"{visit} VISIT"
        elements.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
        elements.append(Spacer(1, 15))

        batches = list(range(0, len(images), 4))

        for idx, i in enumerate(batches):

            batch = images[i:i+4]

            row1, row2 = [], []

            for j, p in enumerate(batch):
                if j < 2:
                    row1.append(resize(p))
                else:
                    row2.append(resize(p))

            while len(row1) < 2: row1.append("")
            while len(row2) < 2: row2.append("")

            t = Table([row1, row2], colWidths=usable_width/2)

            t.setStyle([
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("LEFTPADDING",(0,0),(-1,-1),6),
                ("RIGHTPADDING",(0,0),(-1,-1),6),
                ("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,0),(-1,-1),6),
            ])

            elements.append(t)

            # Spacer ONLY if another batch follows
            if idx != len(batches) - 1:
                elements.append(Spacer(1, 20))

    doc.build(elements)

# ---------------- WORD ---------------- #

def generate_word(visit_images, output_path):

    doc = Document()

    IMG_SIZE = Inches(3)

    first_section = True

    for visit, images in visit_images.items():

        if not images:
            continue

        if not first_section:
            doc.add_page_break()

        first_section = False

        title = visit if visit == "Others" else f"{visit} VISIT"
        doc.add_heading(title, level=1)

        for i in range(0,len(images),4):

            batch = images[i:i+4]

            table = doc.add_table(rows=2, cols=2)
            table.autofit = False

            for row in table.rows:
                for cell in row.cells:
                    cell.width = Inches(3.2)

            idx = 0

            for r in range(2):
                for c in range(2):
                    if idx < len(batch):
                        table.rows[r].cells[c].paragraphs[0].add_run().add_picture(
                            batch[idx],
                            width=IMG_SIZE
                        )
                    idx += 1

    doc.save(output_path)


# ---------------- STREAMLIT UI ---------------- #

st.set_page_config(layout="centered")
st.title("📄 Images → Document Generator")

filename = st.text_input("Enter output filename (without extension)", "visit_report")

output_format = st.radio("Select Output Format", ["PDF","Word","Both"])

visit_uploads = {}

for visit in VISIT_TYPES:
    visit_uploads[visit] = st.file_uploader(
        f"Upload images for {visit}",
        type=["jpg","jpeg","png"],
        accept_multiple_files=True,
        key=visit
    )

if st.button("Generate File"):

    temp_dir = tempfile.mkdtemp()

    visit_images = {}

    total_images = 0

    for visit, files in visit_uploads.items():

        paths = []

        if files:
            for file in files:
                safe_visit = safe_name(visit)
                path = os.path.join(temp_dir, f"{safe_visit}_{file.name}")
                with open(path,"wb") as f:
                    f.write(file.getbuffer())
                paths.append(path)

        visit_images[visit] = paths
        total_images += len(paths)

    if total_images == 0:
        st.error("Please upload at least one image.")
        st.stop()

    pdf_path = os.path.join(temp_dir,f"{filename}.pdf")
    word_path = os.path.join(temp_dir,f"{filename}.docx")

    if output_format in ["PDF","Both"]:
        generate_pdf(visit_images, pdf_path)

    if output_format in ["Word","Both"]:
        generate_word(visit_images, word_path)

    st.success("Document Generated Successfully ✅")

    if output_format in ["PDF","Both"]:
        with open(pdf_path,"rb") as f:
            st.download_button("⬇ Download PDF",f,file_name=f"{filename}.pdf")

    if output_format in ["Word","Both"]:
        with open(word_path,"rb") as f:
            st.download_button("⬇ Download Word",f,file_name=f"{filename}.docx")
