import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import io
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter


EXCEL_PATH = "data/catalogues.xlsx"
IMAGE_FOLDER = "images/"
EXPORT_FOLDER = "exports/"
LOGO_PATH = "assets/logo.png"

os.makedirs(IMAGE_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)


# ============================================================
# DARK MODE
# ============================================================
def apply_dark_mode():
    dark_css = """
    <style>
        body, .stApp {
            background-color: #ADD8E6 !important;
            color: #EEE !important;
        }
        .stButton>button {
            background-color: #444 !important;
            color: white !important;
            border-radius: 8px;
        }
        .stSelectbox, .stNumberInput, .stTextInput {
            background-color: #ADD8E6 !important;
            color: white !important;
        }
        div[data-testid="stMarkdown"] {
            color: #EEE !important;
        }
        .css-1y4p8pa {
            background-color: #ADD8E6 !important;
        }
        img {
            border-radius: 6px;
        }
    </style>
    """
    st.markdown(dark_css, unsafe_allow_html=True)


#apply_dark_mode()


# ============================================================
# EXTRACT IMAGES FROM EXCEL (Correct Method)
# ============================================================
def extract_excel_images():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    image_map = {}

    for img in ws._images:
        # get anchor position
        col = img.anchor._from.col + 1
        row = img.anchor._from.row + 1

        cell_ref = f"{openpyxl.utils.get_column_letter(col)}{row}"

        img_bytes = img._data()
        pil_img = Image.open(io.BytesIO(img_bytes))
        filename = f"img_{cell_ref}.png"

        pil_img.save(IMAGE_FOLDER + filename)
        image_map[cell_ref] = filename

    return image_map


# ============================================================
# LOAD SHEET + LINK IMAGES TO ROWS
# ============================================================
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_PATH)
    image_map = extract_excel_images()

    df["image_file"] = None

    # Column W contains images → row index + 2 = Excel row
    for idx in df.index:
        excel_row = idx + 2
        cell = f"W{excel_row}"

        if cell in image_map:
            df.at[idx, "image_file"] = image_map[cell]

    return df


df = load_data()


# ============================================================
# UI HEADER
# ============================================================
if os.path.exists(LOGO_PATH):
    st.image(LOGO_PATH, width=220)

st.title("💡 Krislite Lighting Finder — Dark Mode Edition")
st.markdown("---")


# ============================================================
# FILTER INPUTS
# ============================================================
col1, col2, col3 = st.columns(3)
required_power = col1.number_input("Power (W)", 0, 200, 10)
required_lumen = col2.number_input("Lumen", 0, 10000, 1000)
required_cri = col3.number_input("Minimum CRI", 0, 100, 80)

mounting = st.selectbox("Mounting", ["Any"] + sorted(df["mounting"].dropna().unique()))
type_option = st.selectbox("Type", ["Any"] + sorted(df["type"].dropna().unique()))
cct_cols = ["2700K","3000K","3500K","4000K","5000K","6500K"]
selected_cct = st.selectbox("Preferred CCT", ["Any"] + cct_cols)

rgb = st.checkbox("RGB Required")
rgbw = st.checkbox("RGBW Required")

view_mode = st.radio("View Mode", ["List View", "Grid View"])
st.markdown("---")


# ============================================================
# PROCESS RESULTS
# ============================================================
if st.button("Search Fixtures"):

    results = df.copy()

    if mounting != "Any":
        results = results[results["mounting"] == mounting]
    if type_option != "Any":
        results = results[results["type"] == type_option]
    if selected_cct != "Any":
        results = results[results[selected_cct] == "Yes"]
    if rgb:
        results = results[results["RGB"] == "Yes"]
    if rgbw:
        results = results[results["RGBW"] == "Yes"]

    # SORTING
    results["power_diff"] = abs(results["power"] - required_power)
    results["lumen_diff"] = abs(results["lumen"] - required_lumen)

    results = results.sort_values(
        ["power_diff", "lumen_diff", "cri"], ascending=[True, True, False]
    )

    st.subheader(f"Found {len(results)} fixtures")

    # EXPORT BUTTONS
    def export_excel():
        out = EXPORT_FOLDER + "results.xlsx"
        results.to_excel(out, index=False)
        return out

    def export_pdf():
        out = EXPORT_FOLDER + "results.pdf"
        doc = SimpleDocTemplate(out, pagesize=letter)
        story = []
        style = getSampleStyleSheet()["Normal"]

        for _, r in results.iterrows():
            story.append(Paragraph(f"<b>{r['model_name']} - {r['model_no']}</b>", style))
            story.append(Paragraph(f"Brand: {r['brand']}", style))
            story.append(Paragraph(f"Type: {r['type']}", style))
            story.append(Paragraph(f"Mounting: {r['mounting']}", style))
            story.append(Paragraph(f"Power: {r['power']}W", style))
            story.append(Paragraph(f"Lumen: {r['lumen']}", style))
            story.append(Paragraph(f"CRI: {r['cri']}", style))
            story.append(Paragraph(f"Input Voltage: {r['ip_v']}", style))
            story.append(Paragraph(f"IP Rating: {r['ip']}", style))

            if r["image_file"]:
                story.append(RLImage(IMAGE_FOLDER + r["image_file"], width=160, height=60))

            story.append(Spacer(1, 15))

        doc.build(story)
        return out

    colA, colB = st.columns(2)
    if colA.button("⬇ Export Excel"):
        fpath = export_excel()
        with open(fpath, "rb") as f:
            st.download_button("Download Excel File", f, "results.xlsx")

    if colB.button("⬇ Export PDF"):
        fpath = export_pdf()
        with open(fpath, "rb") as f:
            st.download_button("Download PDF File", f, "results.pdf")

    st.markdown("---")

    # ============================================================
    # DISPLAY RESULTS
    # ============================================================
    if view_mode == "List View":
        for _, r in results.iterrows():
            with st.container(border=True):
                left, right = st.columns([1, 2])

                with left:
                    if r["image_file"]:
                        st.image(IMAGE_FOLDER + r["image_file"])
                    else:
                        st.write("(No image)")

                with right:
                    st.markdown(f"### {r['model_name']} — {r['model_no']}")
                    st.write(f"**Brand:** {r['brand']}")
                    st.write(f"**Type:** {r['type']}")
                    st.write(f"**Mounting:** {r['mounting']}")
                    st.write(f"**Description:** {r['description']}")
                    st.write(f"**Power:** {r['power']}W")
                    st.write(f"**Lumen:** {r['lumen']}")
                    st.write(f"**CRI:** {r['cri']}")
                    st.write(f"**Input Voltage:** {r['ip_v']}")
                    st.write(f"**IP Rating:** {r['ip']}")
                    cct_list = [c for c in cct_cols if r[c] == "Yes"]
                    st.write(f"**CCT:** {', '.join(cct_list)}")
                    st.write(f"**RGB:** {r['RGB']}")
                    st.write(f"**RGBW:** {r['RGBW']}")
                    st.write(f"**Beam:** {r['beam']}")
                    st.write(f"**Comment:** {r['comment']}")

    else:  # GRID VIEW
        cols = st.columns(4)
        index = 0
        for _, r in results.iterrows():
            with cols[index]:
                with st.container(border=True):
                    if r["image_file"]:
                        st.image(IMAGE_FOLDER + r["image_file"])
                    st.write(f"**{r['model_name']}**")
                    st.write(r["model_no"])
                    st.write(f"{r['power']}W | {r['lumen']}lm")
                    st.write(f"CRI {r['cri']} | IP{r['ip']}")

            index = (index + 1) % 4
