import mimetypes
mimetypes.init()
mimetypes.add_type(
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsx'
)

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
from tempfile import NamedTemporaryFile
from PIL import Image as PILImage
from io import BytesIO
import tempfile
import os
import base64

st.set_page_config(page_title="Servando Recipe Card Generator", layout="wide")

# ── Encode logo ──────────────────────────────────────────────
def get_logo_b64():
    logo_path = os.path.join(os.path.dirname(__file__), "Servando_Branding.jpg")
    try:
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return ""

logo_b64 = get_logo_b64()

# ── CSS ──────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@300;400;500;600&display=swap');

*, *::before, *::after {{ box-sizing: border-box; }}
html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    background-color: #0A1208;
    color: #D4D0C8;
}}
.stApp {{
    background-color: #0A1208;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='900' height='700' viewBox='0 0 900 700'%3E%3Crect width='900' height='700' fill='%230A1208'/%3E%3Cg opacity='0.28'%3E%3Cpath d='M-10,700 C30,580 110,510 165,415 C192,362 180,298 155,258 C198,302 214,370 192,425 C242,358 252,265 230,196 C272,275 260,378 228,445 C276,368 298,254 286,162 C324,250 300,376 262,458 C235,512 196,596 168,700Z' fill='%231A4A1F'/%3E%3Cline x1='155' y1='258' x2='98' y2='700' stroke='%232A6B2F' stroke-width='1.5' opacity='0.45'/%3E%3Cline x1='155' y1='258' x2='188' y2='362' stroke='%232A6B2F' stroke-width='0.9' opacity='0.35'/%3E%3Cline x1='155' y1='258' x2='214' y2='400' stroke='%232A6B2F' stroke-width='0.9' opacity='0.35'/%3E%3Cline x1='155' y1='258' x2='142' y2='480' stroke='%232A6B2F' stroke-width='0.9' opacity='0.35'/%3E%3C/g%3E%3Cg opacity='0.22' transform='translate(860,0)'%3E%3Cpath d='M20,700 C0,580 -42,498 -78,415 C-100,358 -88,288 -64,242 C-98,292 -110,368 -92,425 C-130,352 -136,254 -112,185 C-152,270 -140,386 -106,450 C-140,362 -172,244 -166,138 C-196,242 -168,382 -128,462 C-104,518 -70,600 -46,700Z' fill='%231A4A1F'/%3E%3Cline x1='-64' y1='242' x2='-16' y2='700' stroke='%232A6B2F' stroke-width='1.5' opacity='0.45'/%3E%3C/g%3E%3Cg opacity='0.18' transform='translate(370,-70)'%3E%3Cpath d='M70,0 C26,92 -4,206 12,322 C-28,258 -40,162 -12,80 C-62,126 -74,242 -46,334 C-86,270 -82,160 -56,80 C-104,160 -92,292 -56,390 C-28,430 18,458 64,452 C108,446 150,406 164,350 C182,270 166,154 124,80 C144,160 138,276 112,338 C130,248 126,132 98,58 C78,-16 74,-16 70,0Z' fill='%23163D1A'/%3E%3Cline x1='70' y1='0' x2='46' y2='452' stroke='%232A6B2F' stroke-width='1.2' opacity='0.4'/%3E%3C/g%3E%3Cg opacity='0.14' transform='translate(640,230) rotate(-18)'%3E%3Cpath d='M0,-35 C-28,34 -40,126 -20,202 C-55,148 -63,63 -37,0 C-78,46 -85,144 -58,212 C-90,160 -85,68 -58,0 C-104,70 -92,178 -53,254 C-25,292 14,310 50,304 C88,296 112,258 116,212 C124,138 106,56 66,8 C86,66 84,148 60,204 C76,132 72,58 46,8 C28,-20 10,-25 0,-35Z' fill='%231A4A1F'/%3E%3Cline x1='0' y1='-35' x2='28' y2='304' stroke='%232A6B2F' stroke-width='1.2' opacity='0.4'/%3E%3C/g%3E%3Cg opacity='0.1' transform='translate(150,360) rotate(6)'%3E%3Cpath d='M0,0 C-28,64 -36,150 -18,225 C-52,172 -58,86 -32,22 C-72,64 -78,160 -52,230 C-82,178 -78,82 -52,16 C-94,80 -84,190 -48,264 C-24,306 14,322 48,316 C80,308 104,270 106,226 C114,150 98,66 60,16 C78,70 74,154 50,208 C66,140 62,62 40,16 C24,-10 8,-6 0,0Z' fill='%231A4A1F'/%3E%3C/g%3E%3Cg opacity='0.09' transform='translate(480,400) rotate(-28)'%3E%3Cellipse cx='0' cy='0' rx='20' ry='115' fill='%231A4A1F'/%3E%3Cellipse cx='44' cy='-28' rx='16' ry='92' fill='%231A4A1F' transform='rotate(30)'/%3E%3Cellipse cx='-38' cy='18' rx='16' ry='86' fill='%231A4A1F' transform='rotate(-24)'/%3E%3Cellipse cx='78' cy='-18' rx='14' ry='78' fill='%231A4A1F' transform='rotate(44)'/%3E%3Cellipse cx='-68' cy='28' rx='13' ry='72' fill='%231A4A1F' transform='rotate(-40)'/%3E%3C/g%3E%3Cg opacity='0.06' transform='translate(260,550)'%3E%3Cpath d='M0,-20 C-18,28 -24,90 -12,140 C-36,105 -40,48 -22,8 C-50,32 -54,108 -36,158 C-56,118 -54,52 -36,8 C-66,52 -58,132 -34,182 C-18,205 8,216 34,212 C58,208 74,184 76,158 C82,106 70,46 44,10 C56,48 54,108 38,148 C48,100 46,44 30,10 C18,-8 6,-14 0,-20Z' fill='%231A4A1F'/%3E%3C/g%3E%3C/svg%3E");
    background-attachment: fixed;
    background-size: cover;
    background-position: center;
}}
[data-testid="stSidebar"] {{
    background-color: rgba(8, 14, 8, 0.92) !important;
    border-right: 1px solid #2A3828;
    padding-top: 0 !important;
    backdrop-filter: blur(8px);
}}
[data-testid="stSidebar"] > div:first-child {{ padding-top: 0; }}
.logo-panel {{
    background: linear-gradient(175deg, #1A2E1A 0%, #0E1C0E 60%, #080F08 100%);
    padding: 32px 24px 24px 24px;
    border-bottom: 1px solid #2A3828;
    text-align: center;
    position: relative;
    overflow: hidden;
}}
.logo-panel::before {{
    content: '';
    position: absolute;
    inset: 0;
    background: radial-gradient(ellipse at 50% 0%, rgba(74,107,62,0.18) 0%, transparent 70%);
    pointer-events: none;
}}
.logo-panel img {{
    width: 85%;
    max-width: 220px;
    mix-blend-mode: luminosity;
    filter: brightness(0.85) contrast(1.1) sepia(0.2) hue-rotate(60deg) saturate(0.7);
    border-radius: 4px;
}}
.logo-sub {{
    font-family: 'Inter', sans-serif;
    font-size: 0.6rem;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #5A7A52;
    margin-top: 10px;
}}
.nav-label {{
    font-size: 0.6rem;
    font-weight: 600;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #4A6B3E;
    padding: 20px 0 6px 0;
    border-bottom: 1px solid #1E2E1C;
    margin-bottom: 8px;
}}
.main-header {{
    background: linear-gradient(135deg, rgba(26,46,26,0.85) 0%, rgba(21,37,21,0.85) 100%);
    border: 1px solid #2A3828;
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 20px;
    backdrop-filter: blur(6px);
}}
.main-header-text h2 {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.6rem;
    color: #8CAF7A;
    margin: 0 0 2px 0;
    font-weight: 600;
    letter-spacing: 1px;
}}
.main-header-text p {{
    font-size: 0.7rem;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    color: #4A6B3E;
    margin: 0;
}}
.section-label {{
    font-size: 0.62rem;
    font-weight: 600;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #5A7A52;
    margin: 20px 0 10px 0;
    padding-bottom: 6px;
    border-bottom: 1px solid #1E2E1C;
}}
[data-testid="metric-container"] {{
    background: #111E11 !important;
    border: 1px solid #1E2E1C !important;
    border-radius: 10px !important;
    padding: 14px 16px !important;
}}
[data-testid="metric-container"] label {{
    font-size: 0.62rem !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    color: #5A7A52 !important;
}}
[data-testid="stMetricValue"] {{
    font-family: 'Cormorant Garamond', serif !important;
    font-size: 1.6rem !important;
    font-weight: 600 !important;
    color: #A8C896 !important;
}}
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stTextArea > div > textarea,
.stSelectbox > div > div {{
    background-color: #111E11 !important;
    border: 1px solid #2A3828 !important;
    border-radius: 8px !important;
    color: #D4D0C8 !important;
    font-size: 0.88rem !important;
}}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus,
.stTextArea > div > textarea:focus {{
    border-color: #5A7A52 !important;
    box-shadow: 0 0 0 2px rgba(90,122,82,0.2) !important;
}}
label {{
    color: #8A9E84 !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.3px !important;
}}
.stButton > button {{
    background-color: #2A3E28 !important;
    color: #A8C896 !important;
    border: 1px solid #3A5238 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.8px !important;
    padding: 9px 18px !important;
    transition: all 0.2s !important;
    width: 100%;
}}
.stButton > button:hover {{
    background-color: #3A5238 !important;
    border-color: #5A7A52 !important;
    color: #C8DCC0 !important;
}}
[data-testid="stDownloadButton"] > button {{
    background: linear-gradient(135deg, #3A5238, #2A3E28) !important;
    color: #C8DCC0 !important;
    border: 1px solid #5A7A52 !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    letter-spacing: 1.5px !important;
    text-transform: uppercase !important;
    width: 100% !important;
    padding: 14px !important;
}}
.streamlit-expanderHeader {{
    background-color: #111E11 !important;
    border: 1px solid #1E2E1C !important;
    border-radius: 8px !important;
    color: #A8C896 !important;
    font-size: 0.86rem !important;
}}
.streamlit-expanderContent {{
    background-color: #0E1810 !important;
    border: 1px solid #1E2E1C !important;
    border-top: none !important;
    border-radius: 0 0 8px 8px !important;
}}
[data-testid="stDataFrame"] {{
    border: 1px solid #1E2E1C !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}}
hr {{ border-color: #1E2E1C !important; }}
.stSuccess {{ background-color: #1A2E1A !important; border-color: #3A5238 !important; }}
.stError {{ background-color: #2E1A1A !important; }}
[data-testid="stFileUploader"] {{
    border: 1px dashed #2A3828 !important;
    border-radius: 10px !important;
    background: #111E11 !important;
    padding: 12px !important;
}}
h3 {{
    font-family: 'Cormorant Garamond', serif !important;
    color: #A8C896 !important;
    font-size: 1.3rem !important;
}}
.total-badge {{
    background: linear-gradient(135deg, #1A2E1A, #111E11);
    border: 1px solid #3A5238;
    border-radius: 10px;
    padding: 16px 24px;
    text-align: center;
    margin: 12px 0;
}}
.total-badge .label {{
    font-size: 0.62rem;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #5A7A52;
}}
.total-badge .value {{
    font-family: 'Inter', sans-serif;
    font-size: 1.8rem;
    font-weight: 600;
    color: #A8C896;
    display: block;
    margin-top: 4px;
}}
</style>
""", unsafe_allow_html=True)

# ── Photo helper ─────────────────────────────────────────────
def place_in_box(pil_img, box_w, box_h):
    pil_img = pil_img.convert("RGB")
    pil_img.thumbnail((box_w, box_h), PILImage.LANCZOS)
    canvas = PILImage.new("RGB", (box_w, box_h), (255, 255, 255))
    canvas.paste(pil_img, ((box_w - pil_img.width) // 2,
                            (box_h - pil_img.height) // 2))
    return canvas

# ── Grid layout calculator ───────────────────────────────────
def build_photo_grid(ws, n_cols=4, gap_px=4):
    """
    Read actual column widths (A-H) and row heights from the workbook,
    compute photo box size and exact OneCellAnchor positions for each slot.
    Accounts for print margins so photos stay inside the printed border.
    Returns: (box_w, box_h, anchors) where anchors[i] = (col_idx, col_off_emu, row_height_emu_per_row)
    """
    # ── Column widths (px) for columns A-H ──────────────────
    default_col = ws.sheet_format.defaultColWidth or 8
    col_px = []
    for c in range(1, 9):   # A=1 … H=8
        letter = get_column_letter(c)
        cd = ws.column_dimensions.get(letter)
        units = cd.width if (cd and cd.width) else default_col
        col_px.append(int(units * 7 + 5))

    total_px = sum(col_px)

    # ── Deduct print margins so photos stay inside border ────
    scale    = (ws.page_setup.scale or 100) / 100
    left_px  = int(ws.page_margins.left  * 96 / scale)
    right_px = int(ws.page_margins.right * 96 / scale)
    safe_px  = total_px - left_px - right_px - 10   # 10px safety buffer

    box_w = (safe_px - gap_px * (n_cols - 1)) // n_cols
    box_h = int(box_w * 0.75)   # 4:3

    # Cumulative column positions
    cum = [0]
    for w in col_px:
        cum.append(cum[-1] + w)

    # For each photo column slot, find which Excel column it falls in
    # and the pixel offset from that column's left edge
    col_anchors = []   # list of (excel_col_0based, col_off_px)
    for i in range(n_cols):
        x_px = i * (box_w + gap_px)
        for c in range(len(col_px)):
            if cum[c] <= x_px < cum[c + 1]:
                col_anchors.append((c, x_px - cum[c]))
                break

    # ── Row height in EMU ────────────────────────────────────
    default_row_pt = ws.sheet_format.defaultRowHeight or 15
    PT_TO_EMU = 12700

    def row_emu(r_1based):
        rd = ws.row_dimensions.get(r_1based)
        pt = rd.height if (rd and rd.height) else default_row_pt
        return int(pt * PT_TO_EMU)

    def y_emu_at_row(target_1based):
        return sum(row_emu(r) for r in range(1, target_1based))

    # How many Excel rows does one photo band (box_h + gap) span?
    # We advance the anchor row by this count per photo row
    # so rowOff stays 0 and Excel never clips the image.
    one_row_emu    = int(default_row_pt * PT_TO_EMU)
    band_total_emu = pixels_to_EMU(box_h + gap_px)
    rows_per_band  = max(1, (band_total_emu + one_row_emu - 1) // one_row_emu)

    return box_w, box_h, col_anchors, y_emu_at_row, rows_per_band

# ── SIDEBAR ─────────────────────────────────────────────────
with st.sidebar:
    if logo_b64:
        st.markdown(f"""
        <div class="logo-panel">
            <img src="data:image/jpeg;base64,{logo_b64}" alt="Servando Logo"/>
            <div class="logo-sub">Recipe Card System</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="logo-panel">
            <div style="font-family:'Cormorant Garamond',serif;font-size:1.8rem;color:#8CAF7A;font-weight:700;">SERVANDO</div>
            <div class="logo-sub">Recipe Card System</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="nav-label">📌 Recipe Details</div>', unsafe_allow_html=True)
    recipe_name = st.text_input("Recipe Name")
    category    = st.text_input("Category")

    st.markdown('<div class="nav-label">⚖️ Yield & Serving</div>', unsafe_allow_html=True)
    total_yield  = st.number_input("Total Recipe Yield", min_value=0.0, value=None, placeholder="0", step=1.0, format="%g")
    serving_size = st.number_input("Serving Size",        min_value=0.0, value=None, placeholder="0", step=1.0, format="%g")
    total_yield  = total_yield  or 0.0
    serving_size = serving_size or 0.0
    servings = total_yield / serving_size if serving_size > 0 else 0
    st.metric("No. of Servings", f"{servings:.0f}")

    st.markdown('<div class="nav-label">✍️ Sign-Off</div>', unsafe_allow_html=True)
    prepared_by = st.text_input("Prepared By")
    checked_by  = st.text_input("Checked By")

# ── LOAD COSTS ───────────────────────────────────────────────
cost_df = pd.read_excel("costs.xlsx")
cost_df.columns = cost_df.columns.str.strip()
ingredients = cost_df.iloc[:, 0].dropna().tolist()

if "items" not in st.session_state:
    st.session_state["items"] = []

# ── MAIN HEADER ──────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <div class="main-header-text">
        <h2>Recipe Card Generator</h2>
        <p>Costing &nbsp;·&nbsp; Standardization &nbsp;·&nbsp; Production</p>
    </div>
</div>
""", unsafe_allow_html=True)

left, right = st.columns([1.1, 1], gap="large")

# ════════════════ LEFT COLUMN ════════════════
with left:
    st.markdown('<div class="section-label">➕ Add Ingredient</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        ingredient = st.selectbox("Ingredient", ingredients)
    with col2:
        qty = st.number_input("Quantity", min_value=0.0, value=None, placeholder="0", step=1.0, format="%g", key="qty_input")
        qty = qty or 0.0

    notes_input = st.text_input("Notes", placeholder="e.g. sifted, room temp, finely chopped…")

    if st.button("＋ Add Ingredient"):
        try:
            row = cost_df[cost_df.iloc[:, 0] == ingredient].iloc[0]
            unit_cost = float(row.iloc[1])
            uom       = str(row.iloc[2])
            packaging = 1000 if uom.lower() in ["g", "ml"] else 1
            st.session_state["items"].append({
                "ingredient": ingredient,
                "qty":        qty,
                "packaging":  packaging,
                "uom":        uom,
                "unit_cost":  unit_cost,
                "notes":      notes_input
            })
            st.success(f"Added: {ingredient}")
        except:
            st.error("Error adding ingredient. Check costs.xlsx format.")

    st.markdown('<div class="section-label">📋 Ingredients List</div>', unsafe_allow_html=True)

    delete_index = None
    for i, item in enumerate(st.session_state["items"]):
        with st.expander(f"{item['ingredient']}  —  {item['qty']} {item['uom']}"):
            c1, c2 = st.columns(2)
            with c1:
                item["qty"]       = st.number_input("Qty",       value=float(item["qty"]),       key=f"qty_{i}")
                item["packaging"] = st.number_input("Packaging", value=float(item["packaging"]), key=f"pack_{i}")
            with c2:
                item["uom"]       = st.text_input("UOM",       value=item["uom"],              key=f"uom_{i}")
                item["unit_cost"] = st.number_input("Unit Cost", value=float(item["unit_cost"]), key=f"cost_{i}")
            item["notes"] = st.text_input("Notes", value=item.get("notes",""), key=f"notes_{i}", placeholder="e.g. sifted…")
            if st.button(f"Remove", key=f"delete_{i}"):
                delete_index = i

    if delete_index is not None:
        st.session_state["items"].pop(delete_index)
        st.rerun()

    c_undo, c_clear = st.columns(2)
    with c_undo:
        if st.button("↩ Undo Last") and st.session_state["items"]:
            removed = st.session_state["items"].pop()
            st.toast(f"Removed: {removed['ingredient']}")
            st.rerun()
    with c_clear:
        if st.button("✕ Clear All"):
            st.session_state["items"] = []
            st.rerun()

    st.markdown('<div class="section-label">🧑‍🍳 Procedure</div>', unsafe_allow_html=True)

    if "procedure_steps" not in st.session_state:
        st.session_state["procedure_steps"] = [""]

    for idx in range(len(st.session_state["procedure_steps"])):
        st.session_state["procedure_steps"][idx] = st.text_input(
            f"Step {idx + 1}",
            value=st.session_state["procedure_steps"][idx],
            key=f"step_{idx}",
            placeholder="Describe this step…"
        )

    cp1, cp2 = st.columns(2)
    with cp1:
        if st.button("＋ Add Step"):
            st.session_state["procedure_steps"].append("")
            st.rerun()
    with cp2:
        if st.button("− Remove Last Step") and len(st.session_state["procedure_steps"]) > 1:
            st.session_state["procedure_steps"].pop()
            st.rerun()

    st.markdown('<div class="section-label">📷 Step Photos</div>', unsafe_allow_html=True)
    st.caption("Upload one photo at a time. They will be placed in sequence on the recipe card.")

    if "photo_slots" not in st.session_state:
        st.session_state["photo_slots"] = 1

    images = []
    for idx in range(st.session_state["photo_slots"]):
        uploaded = st.file_uploader(f"Photo {idx + 1}", type=["png", "jpg", "jpeg"], key=f"photo_slot_{idx}")
        images.append(uploaded)

    ph1, ph2 = st.columns(2)
    with ph1:
        if st.button("＋ Add Photo Slot"):
            st.session_state["photo_slots"] += 1
            st.rerun()
    with ph2:
        if st.button("− Remove Last Slot") and st.session_state["photo_slots"] > 1:
            st.session_state["photo_slots"] -= 1
            st.rerun()

# ════════════════ RIGHT COLUMN ════════════════
with right:
    total = 0.0
    srp   = 0.0

    if st.session_state["items"]:
        st.markdown('<div class="section-label">📊 Cost Breakdown</div>', unsafe_allow_html=True)
        df = pd.DataFrame(st.session_state["items"])
        df["Total Cost"] = df["qty"] * df["unit_cost"]
        display_df = df[["ingredient","qty","uom","unit_cost","Total Cost"]].copy()
        display_df.columns = ["Ingredient","Qty","UOM","Unit Cost","Total Cost"]
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        total = df["Total Cost"].sum()
        st.markdown(f"""
        <div class="total-badge">
            <div class="label">Total Recipe Cost</div>
            <span class="value">₱{total:,.2f}</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="section-label">💰 Pricing</div>', unsafe_allow_html=True)
        srp = st.number_input("Selling Price (SRP)", min_value=0.0)

        if srp > 0:
            fc_pct = (total / srp) * 100
            m1, m2 = st.columns(2)
            with m1:
                st.metric("Food Cost %", f"{fc_pct:.1f}%")
            with m2:
                st.metric("Gross Profit", f"₱{srp - total:,.2f}")
    else:
        st.markdown("""
        <div style="text-align:center;padding:60px 20px;color:#3A5238;border:1px dashed #1E2E1C;border-radius:12px;margin-top:40px;">
            <div style="font-size:2rem;margin-bottom:12px;">🍽</div>
            <div style="font-family:'Cormorant Garamond',serif;font-size:1.1rem;color:#5A7A52;">No ingredients added yet</div>
            <div style="font-size:0.75rem;letter-spacing:1px;margin-top:6px;color:#3A5238;">Use the form on the left to begin</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="section-label">📥 Export</div>', unsafe_allow_html=True)

    if st.button("⬇  Generate Recipe Card"):
        try:
            wb = load_workbook("template.xlsx")
            ws = wb.active

            ws["A3"]  = recipe_name
            ws["A6"]  = category
            ws["A55"] = recipe_name
            ws["A58"] = category
            ws["A8"]  = total_yield
            ws["C8"]  = serving_size
            ws["G48"] = srp
            ws["H47"] = prepared_by
            ws["H51"] = checked_by

            start_row = 13
            for i, item in enumerate(st.session_state["items"]):
                r = start_row + i
                ws[f"A{r}"] = item["ingredient"]
                ws[f"B{r}"] = item["qty"]
                ws[f"C{r}"] = item["packaging"]
                ws[f"D{r}"] = item["uom"]
                ws[f"F{r}"] = item["unit_cost"]
                ws[f"H{r}"] = item.get("notes", "")

            for r in range(start_row + len(st.session_state["items"]), 41):
                for col in ["A","B","C","D","E","F","G","H"]:
                    ws[f"{col}{r}"] = None

            row_cursor = 62
            for i, line in enumerate(st.session_state["procedure_steps"], start=1):
                if line.strip():
                    ws[f"A{row_cursor}"] = f"Step {i}: {line}"
                    row_cursor += 1

            # ── PHOTOS: auto-fit to actual sheet width ────────
            #
            # build_photo_grid() reads the real column widths and
            # row heights from the loaded workbook, then computes:
            #   - box_w / box_h: photo cell size that fills the page
            #   - col_anchors:   which Excel column each slot falls
            #                    in, and the pixel offset from its edge
            #   - y_emu_at_row:  exact Y in EMU for any row number
            #
            # OneCellAnchor(col, colOff, row, rowOff) then places
            # each image at the correct cell + sub-cell offset.

            N_COLS  = 4
            GAP_PX  = 4

            box_w, box_h, col_anchors, y_emu_at_row, rows_per_band = build_photo_grid(
                ws, n_cols=N_COLS, gap_px=GAP_PX
            )

            w_emu = pixels_to_EMU(box_w)
            h_emu = pixels_to_EMU(box_h)

            # Photos start 2 rows below the last procedure step
            photo_start_row = row_cursor + 2   # 1-based

            temp_images = []
            valid_images = [img for img in images if img is not None]

            for idx, img_file in enumerate(valid_images):
                try:
                    row_num = idx // N_COLS   # which photo row (0, 1, 2 …)
                    col_num = idx % N_COLS    # which column slot (0-3)

                    excel_col_0, col_off_px = col_anchors[col_num]
                    col_off_emu = pixels_to_EMU(col_off_px)

                    # Advance the anchor row by rows_per_band per photo row
                    # so images never stack on top of each other
                    anchor_row_0based = (photo_start_row - 1) + row_num * rows_per_band

                    img_file.seek(0)
                    fitted = place_in_box(PILImage.open(img_file), box_w, box_h)

                    buf = BytesIO()
                    fitted.save(buf, format="PNG")
                    buf.seek(0)

                    with NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                        tmp.write(buf.read())
                        tp = tmp.name
                    temp_images.append(tp)

                    xl_img        = XLImage(tp)
                    xl_img.width  = box_w
                    xl_img.height = box_h
                    xl_img.anchor = OneCellAnchor(
                        _from=AnchorMarker(
                            col    = excel_col_0,
                            colOff = col_off_emu,
                            row    = anchor_row_0based,
                            rowOff = 0
                        ),
                        ext=XDRPositiveSize2D(w_emu, h_emu)
                    )
                    ws.add_image(xl_img)

                except Exception as img_err:
                    st.warning(f"Could not place photo {idx + 1}: {img_err}")

            # ── Save & download ────────────────────────────────
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                temp_path = tmp.name
            wb.save(temp_path)

            for p in temp_images:
                try: os.remove(p)
                except: pass

            with open(temp_path, "rb") as f:
                file_data = f.read()
            os.remove(temp_path)

            file_name = f"{recipe_name.strip().replace(' ','_')}.xlsx" if recipe_name else "recipe.xlsx"

            st.download_button(
                label="📥  Download Recipe Card",
                data=file_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Failed to generate: {e}")

# ── FOOTER ───────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;padding:32px 0 16px 0;">
    <span style="font-size:0.68rem;letter-spacing:2px;text-transform:uppercase;color:#2A3E28;">
        developed by Dong Cruz
    </span>
</div>
""", unsafe_allow_html=True)