import streamlit as st
import os
import tempfile
from docx import Document
from docx.shared import Inches

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, Image, Spacer, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

from pypdf import PdfReader, PdfWriter

st.markdown("""
<style>

/* Make all labels bold */
label, .stMarkdown p {
    font-weight: 700 !important;
    margin-bottom: 4px !important;
}

/* Reduce space ABOVE inputs */
div[data-testid="stTextInput"],
div[data-testid="stFileUploader"],
div[data-testid="stRadio"] {
    margin-top: -10px !important;
}

/* Reduce space BELOW markdown labels */
.stMarkdown {
    margin-bottom: -10px !important;
}

/* Optional: tighten uploader boxes */
section[data-testid="stFileUploaderDropzone"] {
    padding-top: 10px !important;
    padding-bottom: 10px !important;
}

/* Make TAB text larger */
button[data-baseweb="tab"] {
    font-size: 20px !important;
    font-weight: 700 !important;
    padding: 12px 24px !important;
}

/* Increase tab height */
div[data-baseweb="tab-list"] {
    gap: 12px;
}

/* Active tab underline thicker */
button[data-baseweb="tab"][aria-selected="true"] {
    border-bottom: 3px solid #ff4b4b !important;
}

/* Optional: icons inside tabs bigger */
button[data-baseweb="tab"] span {
    font-size: 20px !important;
}

</style>
""", unsafe_allow_html=True)

PAGE_WIDTH, PAGE_HEIGHT = A4

VISIT_TYPES = [
    "RESIDENCE",
    "EMPLOYMENT / BUSINESS",
    "RESI CUM OFFICE",
    "OTHERS"
]

PDF_CATEGORIES = [
    "Main File",
    "Income Tax Return",
    "Form 26AS",
    "Quotation",
    "Shop Act License",
    "Uddyam Certificate",
    "Others",
    "Images Report"
]

# ---------------- UTILS ---------------- #

def safe_name(text):
    return text.replace("/", "_").replace(" ", "_")

# ---------------- TAB 1 PDF ---------------- #

def generate_pdf(visit_images, output_path):

    LEFT = RIGHT = TOP = 40
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

    first = True

    for visit, images in visit_images.items():

        if not images:
            continue

        if not first:
            elements.append(PageBreak())

        first = False

        title = visit if visit == "Others" else f"{visit} VISIT"
        elements.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
        elements.append(Spacer(1,15))

        batches = list(range(0,len(images),4))

        for idx,i in enumerate(batches):

            batch = images[i:i+4]
            row1,row2=[],[]

            for j,p in enumerate(batch):
                (row1 if j<2 else row2).append(resize(p))

            while len(row1)<2: row1.append("")
            while len(row2)<2: row2.append("")

            t = Table([row1,row2],colWidths=usable_width/2)
            elements.append(t)

            if idx != len(batches)-1:
                elements.append(Spacer(1,20))

    doc.build(elements)

# ---------------- TAB 1 WORD ---------------- #

def generate_word(visit_images, output_path):

    doc = Document()
    IMG_SIZE = Inches(3)
    first = True

    for visit, images in visit_images.items():

        if not images:
            continue

        if not first:
            doc.add_page_break()

        first = False

        title = visit if visit=="Others" else f"{visit} VISIT"
        doc.add_heading(title,level=1)

        for i in range(0,len(images),4):

            batch = images[i:i+4]
            table = doc.add_table(rows=2,cols=2)

            idx=0
            for r in range(2):
                for c in range(2):
                    if idx<len(batch):
                        table.rows[r].cells[c].paragraphs[0].add_run().add_picture(
                            batch[idx], width=IMG_SIZE
                        )
                    idx+=1

    doc.save(output_path)

# ---------------- TAB 2 PDF MERGE ---------------- #

def merge_pdfs(category_files, output_path):

    writer = PdfWriter()

    for category, files in category_files.items():

        for f in files:

            reader = PdfReader(f, strict=False)

            if category == "Form 26AS":
                writer.add_page(reader.pages[0])
            else:
                for page in reader.pages:
                    writer.add_page(page)

    with open(output_path, "wb") as out:
        writer.write(out)

# ---------------- STREAMLIT ---------------- #

st.set_page_config(layout="centered")
st.title("📄 Document Generator")

tab1, tab2 = st.tabs(["🖼 Images Report", "📑 PDF Combiner"])

# ================= TAB 1 ================= #

with tab1:

    st.markdown("**Output filename (without extension)**")
    filename = st.text_input("", "visit_report")

    st.markdown("**Output Format**")
    output_format = st.radio("",["PDF","Word","Both"])


    visit_uploads = {}

    for visit in VISIT_TYPES:
        st.markdown(f"**Upload images for {visit}**")
        visit_uploads[visit] = st.file_uploader(
            "",
            type=["jpg","jpeg","png"],
            accept_multiple_files=True,
            key=visit
        )

    if st.button("Generate Images Report"):

        temp_dir = tempfile.mkdtemp()
        visit_images={}
        total=0

        for visit,files in visit_uploads.items():

            paths=[]

            if files:
                for file in files:
                    safe = safe_name(visit)
                    path = os.path.join(temp_dir,f"{safe}_{file.name}")
                    with open(path,"wb") as f:
                        f.write(file.getbuffer())
                    paths.append(path)

            visit_images[visit]=paths
            total+=len(paths)

        if total==0:
            st.error("Upload at least one image")
            st.stop()

        pdf_path=os.path.join(temp_dir,f"{filename}.pdf")
        word_path=os.path.join(temp_dir,f"{filename}.docx")

        if output_format in ["PDF","Both"]:
            generate_pdf(visit_images,pdf_path)

        if output_format in ["Word","Both"]:
            generate_word(visit_images,word_path)

        st.success("Generated ✅")

        if output_format in ["PDF","Both"]:
            st.download_button("⬇ Download PDF",open(pdf_path,"rb"),file_name=f"{filename}.pdf")

        if output_format in ["Word","Both"]:
            st.download_button("⬇ Download Word",open(word_path,"rb"),file_name=f"{filename}.docx")

# ================= TAB 2 ================= #

with tab2:

    st.markdown("**Combined PDF filename**")
    pdf_filename = st.text_input("", "combined_documents")


    uploads = {}

    for cat in PDF_CATEGORIES:
        st.markdown(f"**{cat}**")
        uploads[cat] = st.file_uploader(
            "",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"pdf_{cat}"
        )

    if st.button("Generate Combined PDF"):

        temp_dir=tempfile.mkdtemp()
        category_files={}
        total=0

        for cat,files in uploads.items():

            paths=[]

            if files:
                for f in files:
                    path=os.path.join(temp_dir,f"{safe_name(cat)}_{f.name}")
                    with open(path,"wb") as out:
                        out.write(f.getbuffer())
                    paths.append(path)

            category_files[cat]=paths
            total+=len(paths)

        if total==0:
            st.error("Upload at least one PDF")
            st.stop()

        output=os.path.join(temp_dir,f"{pdf_filename}.pdf")

        merge_pdfs(category_files,output)

        st.success("Combined PDF Ready ✅")

        st.download_button("⬇ Download Combined PDF",open(output,"rb"),file_name=f"{pdf_filename}.pdf")
