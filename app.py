import streamlit as st
import os
import tempfile
from docx import Document
from docx.shared import Inches

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, Image, Spacer
from reportlab.lib import colors

from docx import Document
from docx.shared import Inches
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak

PAGE_WIDTH, PAGE_HEIGHT = A4


# ---------- IMAGE RESIZE (PDF) ---------- #

def get_resized_image(path):
    img = Image(path)

    max_width = PAGE_WIDTH / 2 - 40
    max_height = PAGE_HEIGHT / 2 - 40

    ratio = min(
        max_width / img.drawWidth,
        max_height / img.drawHeight
    )

    img.drawWidth *= ratio
    img.drawHeight *= ratio
    img.hAlign = "CENTER"

    return img


# ---------- PDF GENERATION ---------- #

def generate_pdf(image_paths, visit_type, output_path):

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

    elements.append(Paragraph(f"<b>Visit Type: {visit_type}</b>", styles["Title"]))
    elements.append(Spacer(1,15))

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

    for i in range(0,len(image_paths),4):

        batch = image_paths[i:i+4]

        row1,row2=[],[]

        for j,p in enumerate(batch):
            if j<2:
                row1.append(resize(p))
            else:
                row2.append(resize(p))

        while len(row1)<2: row1.append("")
        while len(row2)<2: row2.append("")

        t = Table([row1,row2],colWidths=usable_width/2)

        t.setStyle([
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),6),
            ("RIGHTPADDING",(0,0),(-1,-1),6),
            ("TOPPADDING",(0,0),(-1,-1),6),
            ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ])

        elements.append(t)
        elements.append(PageBreak())

    doc.build(elements)


# ---------- WORD GENERATION ---------- #

def generate_word(image_paths, visit_type, output_path):

    doc = Document()

    doc.add_heading(f"Visit Type: {visit_type}", level=1)

    IMG_SIZE = Inches(3)

    for i in range(0, len(image_paths), 4):

        batch = image_paths[i:i+4]

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

        doc.add_page_break()

    doc.save(output_path)

# ---------- STREAMLIT UI ---------- #

st.set_page_config(layout="centered")
st.title("📄 Images → Document Generator")

visit_type = st.selectbox(
    "Select Visit Type",
    [
        "RESIDENCE VISIT",
        "EMPLOYMENT / BUSINESS VISIT",
        "RESI CUM OFFICE VISIT"
    ]
)

output_format = st.radio(
    "Select Output Format",
    ["PDF", "Word", "Both"]
)

uploaded_files = st.file_uploader(
    "Upload Images",
    type=["jpg","jpeg","png"],
    accept_multiple_files=True
)

if uploaded_files:

    st.success(f"{len(uploaded_files)} images uploaded")

    temp_dir = tempfile.mkdtemp()
    image_paths = []

    for file in uploaded_files:
        path = os.path.join(temp_dir, file.name)
        with open(path,"wb") as f:
            f.write(file.getbuffer())
        image_paths.append(path)

    if st.button("Generate File"):

        pdf_path = os.path.join(temp_dir,"output.pdf")
        word_path = os.path.join(temp_dir,"output.docx")

        if output_format in ["PDF","Both"]:
            generate_pdf(image_paths, visit_type, pdf_path)

        if output_format in ["Word","Both"]:
            generate_word(image_paths, visit_type, word_path)

        st.success("Document Generated Successfully ✅")

        if output_format in ["PDF","Both"]:
            with open(pdf_path,"rb") as f:
                st.download_button("⬇ Download PDF",f,file_name="images.pdf")

        if output_format in ["Word","Both"]:
            with open(word_path,"rb") as f:
                st.download_button("⬇ Download Word",f,file_name="images.docx")
