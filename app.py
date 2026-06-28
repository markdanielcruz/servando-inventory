import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import pandas as pd
import json
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Servando Inventory",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@300;400;500;600&display=swap');

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #0A1208;
    color: #D4D0C8;
}
.stApp {
    background-color: #0A1208;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='900' height='700' viewBox='0 0 900 700'%3E%3Crect width='900' height='700' fill='%230A1208'/%3E%3Cg opacity='0.28'%3E%3Cpath d='M-10,700 C30,580 110,510 165,415 C192,362 180,298 155,258 C198,302 214,370 192,425 C242,358 252,265 230,196 C272,275 260,378 228,445 C276,368 298,254 286,162 C324,250 300,376 262,458 C235,512 196,596 168,700Z' fill='%231A4A1F'/%3E%3Cline x1='155' y1='258' x2='98' y2='700' stroke='%232A6B2F' stroke-width='1.5' opacity='0.45'/%3E%3C/g%3E%3Cg opacity='0.22' transform='translate(860,0)'%3E%3Cpath d='M20,700 C0,580 -42,498 -78,415 C-100,358 -88,288 -64,242 C-98,292 -110,368 -92,425 C-130,352 -136,254 -112,185 C-152,270 -140,386 -106,450 C-140,362 -172,244 -166,138 C-196,242 -168,382 -128,462 C-104,518 -70,600 -46,700Z' fill='%231A4A1F'/%3E%3C/g%3E%3Cg opacity='0.18' transform='translate(370,-70)'%3E%3Cpath d='M70,0 C26,92 -4,206 12,322 C-28,258 -40,162 -12,80 C-62,126 -74,242 -46,334 C-86,270 -82,160 -56,80 C-104,160 -92,292 -56,390 C-28,430 18,458 64,452 C108,446 150,406 164,350 C182,270 166,154 124,80 C144,160 138,276 112,338 C130,248 126,132 98,58 C78,-16 74,-16 70,0Z' fill='%23163D1A'/%3E%3C/g%3E%3C/svg%3E");
    background-attachment: fixed;
    background-size: cover;
    background-position: center;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: rgba(8, 14, 8, 0.92) !important;
    border-right: 1px solid #2A3828;
    backdrop-filter: blur(8px);
}
[data-testid="stSidebar"] * { color: #D4D0C8 !important; }

/* Header */
.main-header {
    background: linear-gradient(135deg, rgba(26,46,26,0.85) 0%, rgba(21,37,21,0.85) 100%);
    border: 1px solid #2A3828;
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 20px;
    backdrop-filter: blur(6px);
}
.main-header h1 {
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.8rem;
    color: #8CAF7A;
    margin: 0 0 2px 0;
    font-weight: 600;
    letter-spacing: 1px;
}
.main-header p {
    font-size: 0.7rem;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    color: #4A6B3E;
    margin: 0;
}

/* Cards */
.card {
    background: #111E11;
    border: 1px solid #1E2E1C;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}

/* Buttons */
.stButton > button {
    background-color: #2A3E28 !important;
    color: #A8C896 !important;
    border: 1px solid #3A5238 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.8px !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background-color: #3A5238 !important;
    border-color: #5A7A52 !important;
    color: #C8DCC0 !important;
}

/* Input fields */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stTextArea > div > textarea,
.stSelectbox > div > div {
    background-color: #111E11 !important;
    border: 1px solid #2A3828 !important;
    border-radius: 8px !important;
    color: #D4D0C8 !important;
    font-size: 0.88rem !important;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: #5A7A52 !important;
    box-shadow: 0 0 0 2px rgba(90,122,82,0.2) !important;
}

label {
    color: #8A9E84 !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
}

/* Section titles */
.section-title {
    font-size: 0.62rem;
    font-weight: 600;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #5A7A52;
    margin-bottom: 1rem;
    padding-bottom: 6px;
    border-bottom: 1px solid #1E2E1C;
    display: block;
}

/* Log entry */
.log-entry {
    background: #0E1810;
    border-left: 3px solid #3A5238;
    padding: 0.75rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 0.5rem;
    font-size: 0.9rem;
    color: #D4D0C8;
}
.log-entry .timestamp { color: #4A6B3E; font-size: 0.8rem; }

/* PO item row */
.po-item-row {
    background: #0E1810;
    border: 1px solid #2A3828;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    margin-bottom: 0.5rem;
    color: #D4D0C8;
}

/* Badge */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.badge-beverage { background: #0E1E2E; color: #6AABCC; }
.badge-dry { background: #1E1A0E; color: #C4A840; }
.badge-fresh { background: #0E1E12; color: #5AAA6A; }
.badge-wet { background: #0E1C1A; color: #4AAA96; }
.badge-beef { background: #1E0E0E; color: #CC6A6A; }
.badge-chicken { background: #1E160E; color: #CC8A4A; }
.badge-seafood { background: #0E1620; color: #4A8ACC; }
.badge-pork { background: #1E0E0E; color: #CC5A5A; }
.badge-rtc { background: #1A0E1E; color: #AA6ACC; }

/* Success/info boxes */
.success-box {
    background: #1A2E1A;
    border: 1px solid #3A5238;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    color: #A8C896;
    margin: 0.5rem 0;
}
.info-box {
    background: #111E11;
    border: 1px solid #2A3828;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    color: #8CAF7A;
    margin: 0.5rem 0;
}

/* Metrics */
[data-testid="metric-container"] {
    background: #111E11 !important;
    border: 1px solid #1E2E1C !important;
    border-radius: 10px !important;
    padding: 14px 16px !important;
}
[data-testid="metric-container"] label {
    font-size: 0.62rem !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    color: #5A7A52 !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Cormorant Garamond', serif !important;
    font-size: 1.6rem !important;
    font-weight: 600 !important;
    color: #A8C896 !important;
}

/* Table */
[data-testid="stDataFrame"] {
    border: 1px solid #1E2E1C !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}
hr { border-color: #1E2E1C !important; }
h3 {
    font-family: 'Cormorant Garamond', serif !important;
    color: #A8C896 !important;
    font-size: 1.3rem !important;
}
</style>
""", unsafe_allow_html=True)

# ── Google Sheets Connection ───────────────────────────────────────────────────
SPREADSHEET_ID = "1pYqecDqe_qUwkexRiG5mmOaYV58_1uoRCx_Yy4CepNQ"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def get_gsheet_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet():
    client = get_gsheet_client()
    return client.open_by_key(SPREADSHEET_ID)

def ensure_sheet(spreadsheet, name, headers=None):
    """Get or create a worksheet by name."""
    try:
        ws = spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=2000, cols=30)
        if headers:
            ws.append_row(headers)
    return ws

# ── Sheet Names & Headers ──────────────────────────────────────────────────────
MASTER_SHEET    = "MASTER LIST"
LOG_SHEET       = "DAILY LOG"

MASTER_HEADERS  = ["ITEM", "UNIT COST", "UNIT OF MEASURE", "CATEGORY", "ACTIVE"]
LOG_HEADERS     = ["TIMESTAMP", "MONTH", "DATE", "ITEM", "STAFF", "PO NUMBER", "DEPARTMENT",
                   "ADD'L/IN", "OVER", "DAMAGED/OUT", "CAFÉ", "BAR",
                   "KITCHEN ALA CARTE", "KITCHEN BANQUET", "RESTO", "NOTES"]
SUMMARY_HEADERS = ["ITEM", "UNIT COST", "UNIT OF MEASURE", "CATEGORY",
                   "BEGINNING STOCKS", "ADD'L/IN", "OVER", "DAMAGED/OUT",
                   "CAFÉ", "BAR", "KITCHEN ALA CARTE", "KITCHEN BANQUET",
                   "RESTO", "ENDING STOCKS", "TOTAL WORTH OF STOCKS"]

DEPARTMENTS = ["Restaurant", "Cafe", "Bar", "Banquet", "Warehouse / General", "Others"]
FIELD_TYPES = ["Add'l / In", "Over", "Damaged / Out", "Café", "Bar", "Kitchen Ala Carte", "Kitchen Banquet", "Resto"]

# ── Data helpers ───────────────────────────────────────────────────────────────
def load_master(spreadsheet):
    ws = ensure_sheet(spreadsheet, MASTER_SHEET, MASTER_HEADERS)
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=MASTER_HEADERS)

def load_log(spreadsheet):
    ws = ensure_sheet(spreadsheet, LOG_SHEET, LOG_HEADERS)
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=LOG_HEADERS)

def load_summary(spreadsheet, month_label):
    sheet_name = f"SUMMARY - {month_label}"
    ws = ensure_sheet(spreadsheet, sheet_name, SUMMARY_HEADERS)
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=SUMMARY_HEADERS), ws

def month_label(y, m):
    return datetime(y, m, 1).strftime("%b %Y").upper()

def current_month_label():
    n = datetime.now()
    return month_label(n.year, n.month)

def generate_po_number(entry_date, department):
    """Generate a PO number like PO-20260628-CAFE-001"""
    dept_code = department.upper().replace(" ", "").replace("/", "")[:4]
    date_str = entry_date.strftime("%Y%m%d")
    ts = datetime.now().strftime("%H%M%S")
    return f"PO-{date_str}-{dept_code}-{ts}"

def update_summary_for_item(spreadsheet, entry_date, selected_item, addin, over, damaged, cafe, bar, kitchen, banquet, resto):
    """Update summary sheet for a single item's quantities."""
    month_lbl = entry_date.strftime("%b %Y").upper()
    sum_sheet = f"SUMMARY - {month_lbl}"
    try:
        sum_ws = spreadsheet.worksheet(sum_sheet)
        records = sum_ws.get_all_records()
        for idx, rec in enumerate(records):
            if rec["ITEM"] == selected_item:
                row_num = idx + 2
                cols = ["ADD'L/IN","OVER","DAMAGED/OUT","CAFÉ","BAR",
                        "KITCHEN ALA CARTE","KITCHEN BANQUET","RESTO"]
                vals  = [addin, over, damaged, cafe, bar, kitchen, banquet, resto]
                for col, val in zip(cols, vals):
                    if val:
                        cur = float(rec.get(col, 0) or 0)
                        col_idx = SUMMARY_HEADERS.index(col) + 1
                        sum_ws.update_cell(row_num, col_idx, cur + val)

                # Recalculate ending & worth
                updated = sum_ws.row_values(row_num)
                def gv(i): return float(updated[i-1]) if updated[i-1] else 0
                ending = gv(5)+gv(6)+gv(7)-gv(8)-gv(9)-gv(10)-gv(11)-gv(12)-gv(13)
                cost_v = gv(2)
                sum_ws.update_cell(row_num, 14, round(ending, 4))
                sum_ws.update_cell(row_num, 15, round(ending * cost_v, 4))
                break
    except Exception:
        pass

def recompute_summary(spreadsheet, month_lbl, master_df, log_df):
    """Recompute summary for a month from the log and write it back."""
    sheet_name = f"SUMMARY - {month_lbl}"
    ws = ensure_sheet(spreadsheet, sheet_name, SUMMARY_HEADERS)
    ws.clear()
    ws.append_row(SUMMARY_HEADERS)

    month_log = log_df[log_df["MONTH"] == month_lbl] if not log_df.empty else pd.DataFrame()

    rows = []
    for _, item in master_df[master_df["ACTIVE"] == "YES"].iterrows():
        name = item["ITEM"]
        ilog = month_log[month_log["ITEM"] == name] if not month_log.empty else pd.DataFrame()

        def col_sum(c):
            if ilog.empty or c not in ilog.columns:
                return 0
            vals = pd.to_numeric(ilog[c], errors="coerce").fillna(0)
            return float(vals.sum())

        addin    = col_sum("ADD'L/IN")
        over     = col_sum("OVER")
        damaged  = col_sum("DAMAGED/OUT")
        cafe     = col_sum("CAFÉ")
        bar      = col_sum("BAR")
        kitchen  = col_sum("KITCHEN ALA CARTE")
        banquet  = col_sum("KITCHEN BANQUET")
        resto    = col_sum("RESTO")

        begin_key = f"BEGIN_{month_lbl}_{name}"
        beginning = st.session_state.get(begin_key, 0)

        ending = beginning + addin + over - damaged - cafe - bar - kitchen - banquet - resto
        cost   = float(item["UNIT COST"]) if item["UNIT COST"] else 0
        worth  = ending * cost

        rows.append([
            name, cost, item["UNIT OF MEASURE"], item["CATEGORY"],
            beginning, addin, over, damaged,
            cafe, bar, kitchen, banquet, resto,
            round(ending, 4), round(worth, 4)
        ])

    if rows:
        ws.append_rows(rows)

# ── Initialize Master List from existing Excel ─────────────────────────────────
INITIAL_ITEMS = [["Absolut Vodka Blue",5100.0,0.92,"ml","beverage"],["Accord Powder",0,0.19,"gram","dry"],["Agave Syrup 1.02kg",14280.0,1.088235294117647,"ml","wet"],["Almond - Ground",3000.0,1.08,"gram","dry"],["Amaretto Disaronno 700 ml",5600.0,1.32,"ml","beverage"],["Ampalaya - Fresh",0,0.18,"gram","fresh"],["Ampalaya - Dahon",0,0.41,"gram","fresh"],["Anchovies",448.0,3.3035714285714284,"gram","wet"],["Anisado Wine",0,0.21428571428571427,"ml","wet"],["Aperol 700 ml",4900.0,1.2857142857142858,"ml","beverage"],["Apog Powder",0,3.01,"gram","dry"],["Apple - Red",0,0.16,"gram","fresh"],["Atsuete / Annatto - Seeds",0,0.85,"gram","dry"],["Atsuete / Annatto - Powdered",0,0.33,"gram","dry"],["Bacardi - Gold Rum 750ml",1500.0,0.9333333333333333,"ml","beverage"],["Bacardi - Premium Black 750ml",3000.0,1.1666666666666667,"ml","beverage"],["Bacardi - Superior White 750 ml",1500.0,2.3066666666666666,"ml","beverage"],["Baguio Beans",0,0.18,"gram","fresh"],["Bailey's Original 700 ml",3500.0,1.1428571428571428,"bottle","beverage"],["Baker's Best Margarine 225g",4050.0,0.28,"gram","wet"],["Baking Powder",1000.0,0.22,"gram","dry"],["Baking Soda",4535.0,0.4852941176470588,"gram","dry"],["Balls - Mozarella",0,1.16,"gram","rtc"],["Balls - Shrimp",0,0.215,"gram","rtc"],["Balls - Squid",0,0.297,"gram","rtc"],["Bambi Spring Roll Wrapper - Big 25 pcs",0,10.0,"piece","rtc"],["Bambi Spring Roll Wrapper Small - 35 pcs",0,8.0,"piece","rtc"],["Banana - Lakatan",0,0.14,"gram","fresh"],["Banana - Saba",0,0.08,"gram","fresh"],["Banana - Sweetened, for Halo halo per gram",13000.0,0.19,"gram","fresh"],["Banana Blossom",0,0.57,"gram","dry"],["Banana Essence",0,3.75,"ml","wet"],["Banana Ketchup",0,0.77,"ml","wet"],["Bangus - Daing (pack of 3, per piece)",0,75.0,"piece","seafood"],["Bangus - Smoked Milkfish, Tinapa (420 grams)",9.0,360.0,"piece","seafood"],["Bangus - Unseasoned Fillet - 520 grams",13.0,298.0,"piece","seafood"],["Barbeque Sauce - Ready made",0,2.2,"ml","wet"],["Barrio Fiesta Bagoong Regular",7000.0,0.45,"gram","wet"],["Basil Leaves - Dried",0,1.49,"gram","dry"],["Bay Leaves",0,3.0,"gram","dry"],["Beans - for Halo halo per gram",12000.0,0.24,"gram","wet"],["Beef - Brisket",78670.0,0.48,"gram","beef"],["Beef - Feet",0,0.24,"gram","beef"],["Beef - Ground",5000.0,0.395,"gram","beef"],["Beef - Shank",40120.09,0.43,"gram","beef"],["Beef - Skin",5000.0,0.22,"gram","beef"],["Belgian Waffle Mix",0,0.328,"gram","dry"],["Bell Pepper - Green",0,0.3,"gram","fresh"],["Bell Pepper - Red",0,0.3,"gram","fresh"],["Black Beans",0,0.47,"gram","wet"],["Black Olives - Pitted, Sliced",6075.0,0.6977777777777778,"gram","wet"],["Black Pepper - Ground",1000.0,0.698,"gram","dry"],["Black Pepper - Whole",0,1.2,"gram","dry"],["Blue Curacao 1 liter",2000.0,0.52,"ml","beverage"],["Blue Lemonade Powdered Juice 500g",2000.0,1.175,"gram","beverage"],["Bombay Sapphire Gin 750 ml",1900.0,1.6,"ml","beverage"],["Bread Crumbs - Fine",0,0.18,"gram","dry"],["Bread Crumbs - Japanese",3000.0,0.18,"gram","dry"],["Broccoli",0,0.12,"gram","fresh"],["Broth Cubes - Beef",2220.0,0.75,"gram","rtc"],["Broth Cubes - Chicken",600.0,0.75,"gram","rtc"],["Broth Cubes - Pork",840.0,0.75,"gram","rtc"],["Broth Cubes - Shrimp",2400.0,0.75,"gram","rtc"],["Butter - Unsalted",0,0.3,"gram","rtc"],["Cabbage",0,0.1,"gram","fresh"],["Cabbage - Red",0,0.42,"gram","fresh"],["Cajun Powder",0,1.7875,"gram","wet"],["Calamansi - Concentrate, with Honey",0,0.48,"ml","wet"],["Calamansi - Fresh",0,0.14,"gram","fresh"],["Calamansi - Pure, Unsweetened",1000.0,0.2,"ml","wet"],["Campari Bitter 750 ml",3000.0,1.56,"ml","beverage"],["Canada Dry Ginger Ale 355 ml",36.0,50.0,"can","beverage"],["Canned Tuna",0,0.361,"can","rtc"],["Capers",400.0,0.9614035087719298,"gram","wet"],["Caramel Sauce - Torani",0,0.65,"ml","wet"],["Carrots",0,0.19,"gram","fresh"],["Casa Okinawa Flavor",0,0.43,"gram","dry"],["Cascina Italian Pesto 540 g per bottle",9.0,394.0,"jar","wet"],["Cauliflower",0,0.32,"gram","fresh"],["Cayenne pepper",450.0,2.8846153846153846,"gram","dry"],["Celery",0,0.18,"gram","fresh"],["Cheese - Sauce (500g)",0,0.65,"ml","wet"],["Cheese - Cheddar - Block",10020.0,0.4025,"gram","rtc"],["Cheese - Filled",2700.0,0.31,"gram","rtc"],["Cheese - Mozarella - Block",38890.0,0.5,"gram","rtc"],["Cheese - Parmesan - Ground",10938.0,1.18,"gram","rtc"],["Cheese - Parmesan Block",0,1.66,"gram","rtc"],["Cheese - Quickmelt",3600.0,0.4633333333333333,"gram","rtc"],["Chicken - Breast Fillet",12000.0,0.26,"gram","chicken"],["Chicken - Breast, Bone In",0,0.24,"gram","chicken"],["Chicken - Leg Quarter Bone In",15000.0,0.21,"gram","chicken"],["Chicken - Leg Quarter Fillet Boneless",56000.0,0.335,"gram","chicken"],["Chicken - Liver",6000.0,0.265,"gram","chicken"],["Chicken - Skin",0,0.175,"gram","chicken"],["Chicken - Whole",0,0.235,"gram","chicken"],["Chicken - Wings",0,0.18,"gram","chicken"],["Chili - Flakes",0,0.55,"gram","dry"],["Chili - Green",0,0.25,"gram","fresh"],["Chili - Labuyo",0,0.2,"gram","fresh"],["Chili - Leaves",0,0.12,"gram","dry"],["Chili - Powder",0,0.27,"gram","dry"],["Chorizo - Bilbao",0,1.9857142857142858,"gram","meat"],["Chorizo - Macao",0,0.6363636363636364,"gram","meat"],["Chorizo - Spanish",0,0.65,"gram","meat"],["Cinnamon - Ground",200.0,0.745,"gram","dry"],["Cinnamon Sticks",539.0,1.39,"gram","dry"],["Cloves - Powder",0,2.744186046511628,"gram","dry"],["Cocoa Powder",1000.0,2.331288343558282,"gram","dry"],["Coconut - Shredded",0,55.0,"gram","fresh"],["Coconut Cream per can",8.0,55.0,"can","wet"],["Coconut Rum - Malibu",5600.0,1.0271428571428571,"ml","beverage"],["Coconut Water - Vita Coco",0,0.15151515151515152,"ml","beverage"],["Coffee Beans - Silcafe",0,0.73,"gram","dry"],["Coke 1.5L per ml",0,0.095,"ml","beverage"],["Coke Original 320 ml",50.0,37.0,"can","beverage"],["Coke Zero 320 ml",67.0,37.0,"can","beverage"],["Cooking Sake1 Liter per ml",0,0.23,"ml","wet"],["Coriander Powder",0,0.49,"gram","dry"],["Corn - Kernels, Canned per gram",425.0,0.14285714285714285,"gram","wet"],["Corn - Sweet Style, Canned per gram",425.0,0.18823529411764706,"gram","wet"],["Corn Flakes",0,0.3073770491803279,"gram","dry"],["Cornstarch",0,0.07,"gram","dry"],["Corona Beer 330 ml",144.0,70.0,"bottle","beverage"],["Crab Meat, Frozen",5000.0,1.327,"gram","seafood"],["Crab Paste",908.0,0.8810572687224669,"gram","wet"],["Crab Stick",0,0.329,"gram","seafood"],["Cream - All Purpose, 250ml per ml",12000.0,0.32,"ml","wet"],["Cream - Coconut per ml",57200.0,0.2125,"ml","wet"],["Cream - Cooking",0,0.35,"ml","wet"],["Cream - Full Whipping",0,0.12,"liter","wet"],["Cream - Mascarpone",11000.0,0.89,"ml","wet"],["Cream - Sour",0,0.496,"ml","wet"],["Cream Cheese",1125.0,0.8666666666666667,"gram","rtc"],["Cream Dory - Fillet",3000.0,0.135,"gram","seafood"],["Cream Of Mushroom, Powdered",1550.0,0.8064516129032258,"gram","wet"],["Cream of Tartar",0,0.962,"gram","dry"],["Creamer Sachet (100 pcs, per piece)",500.0,2.0,"piece","dry"],["Crushed Graham",1000.0,0.225,"gram","dry"],["Cucumber - Fresh",0,0.13,"gram","fresh"],["Cumin - Ground",600.0,0.842,"gram","dry"],["Curacao Triple Sec 750 ml",0,0.6,"ml","beverage"],["Custard Powder - Cremyvit",18000.0,0.51,"gram","dry"]]

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='background: linear-gradient(175deg, #1A2E1A 0%, #0E1C0E 60%, #080F08 100%);
                padding: 28px 20px 20px 20px;
                border-bottom: 1px solid #2A3828;
                text-align: center;
                margin-bottom: 8px;'>
        <div style='font-family: Cormorant Garamond, serif; font-size: 1.6rem; font-weight: 700; color: #8CAF7A; letter-spacing: 2px;'>SERVANDO</div>
        <div style='font-size: 0.58rem; letter-spacing: 3px; text-transform: uppercase; color: #4A6B3E; margin-top: 6px;'>Inventory System</div>
    </div>
    <div style='font-size:0.58rem; font-weight:600; letter-spacing:3px; text-transform:uppercase;
                color:#4A6B3E; padding: 16px 12px 6px 12px; border-bottom: 1px solid #1E2E1C; margin-bottom: 8px;'>
        Navigation
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("", [
        "📋 Purchase Order",
        "📦 Ingredients",
        "📅 Month Manager",
        "🔍 Search & History",
        "📊 Summary",
        "⬇️ Export to Excel",
        "⚙️ Setup"
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown(f"<div style='font-size:0.65rem; letter-spacing:1.5px; text-transform:uppercase; color:#3A5238; padding: 4px 0;'>Today: {date.today().strftime('%B %d, %Y')}</div>", unsafe_allow_html=True)

# ── Load data ──────────────────────────────────────────────────────────────────
try:
    spreadsheet = get_spreadsheet()
    connected = True
except Exception as e:
    connected = False
    st.error(f"❌ Could not connect to Google Sheets: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SETUP
# ══════════════════════════════════════════════════════════════════════════════
if page == "⚙️ Setup":
    st.markdown('<div class="main-header"><div><h1>⚙️ Setup</h1><p>Initialize your inventory system</p></div></div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Initialize Master List from May 2026 Excel</div>', unsafe_allow_html=True)
    st.write("This will populate the Master List with all 556 items from your existing inventory file. Run this **once** when setting up.")

    if st.button("🚀 Import Items from May 2026 Excel", type="primary"):
        ws = ensure_sheet(spreadsheet, MASTER_SHEET, MASTER_HEADERS)
        existing = ws.get_all_records()
        existing_names = {r["ITEM"] for r in existing}

        new_rows = []
        for row in INITIAL_ITEMS:
            name = row[0]
            if name not in existing_names:
                new_rows.append([name, row[2], row[3], row[4], "YES"])

        if new_rows:
            ws.append_rows(new_rows)
            st.success(f"✅ Imported {len(new_rows)} items to Master List!")
        else:
            st.info("ℹ️ All items already exist in Master List.")

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Initialize May 2026 Summary (Beginning Stocks)</div>', unsafe_allow_html=True)
    st.write("This sets up the May 2026 summary with beginning stocks from your Excel file.")

    if st.button("📅 Setup May 2026 Beginning Stocks", type="primary"):
        master_df = load_master(spreadsheet)
        log_df = load_log(spreadsheet)
        sheet_name = "SUMMARY - MAY 2026"
        ws = ensure_sheet(spreadsheet, sheet_name, SUMMARY_HEADERS)
        existing = ws.get_all_records()
        if existing:
            st.info("May 2026 summary already exists.")
        else:
            begin_map = {row[0]: row[1] for row in INITIAL_ITEMS}
            rows = []
            for _, item in master_df[master_df["ACTIVE"] == "YES"].iterrows():
                name = item["ITEM"]
                beginning = begin_map.get(name, 0)
                cost = float(item["UNIT COST"]) if item["UNIT COST"] else 0
                rows.append([name, cost, item["UNIT OF MEASURE"], item["CATEGORY"],
                              beginning, 0, 0, 0, 0, 0, 0, 0, 0, beginning, beginning * cost])
            if rows:
                ws.append_rows(rows)
                st.success(f"✅ May 2026 initialized with {len(rows)} items!")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: PURCHASE ORDER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Purchase Order":
    st.markdown('<div class="main-header"><div><h1>📋 Purchase Order</h1><p>Build a PO with multiple items, then submit all at once</p></div></div>', unsafe_allow_html=True)

    master_df = load_master(spreadsheet)
    if master_df.empty:
        st.warning("⚠️ No ingredients found. Please go to Setup first.")
        st.stop()

    active_items = master_df[master_df["ACTIVE"] == "YES"]["ITEM"].tolist()

    # ── Initialize PO cart in session state ──
    if "po_items" not in st.session_state:
        st.session_state.po_items = []

    # ── PO Header ──────────────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">PO Details</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        entry_date = st.date_input("📅 Date", value=date.today(), key="po_date")
    with col2:
        staff_name = st.text_input("👤 Staff Name", placeholder="Enter your name", key="po_staff")
    with col3:
        department_choice = st.selectbox("🏢 Department", DEPARTMENTS, key="po_dept")

    if department_choice == "Others":
        others_text = st.text_input("✏️ Please specify department", placeholder="e.g. Maintenance, Security...", key="po_others")
        department = others_text.strip() if others_text.strip() else "Others"
    else:
        department = department_choice

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Add Item to PO ─────────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Add Item to PO</div>', unsafe_allow_html=True)

    search_query = st.text_input("🔍 Search Ingredient", placeholder="Type to search... e.g. 'chicken', 'beef'", key="po_search")
    filtered = [i for i in active_items if search_query.lower() in i.lower()] if search_query else active_items

    col_item, col_field, col_qty = st.columns([3, 2, 1])
    with col_item:
        selected_item = st.selectbox("Select Ingredient", filtered, key="po_item_select")
    with col_field:
        field_type = st.selectbox("Movement Type", FIELD_TYPES, key="po_field_type")
    with col_qty:
        quantity = st.number_input("Qty", min_value=0.01, step=0.01, format="%.2f", key="po_qty")

    item_notes = st.text_input("📝 Item Notes (optional)", key="po_item_notes")

    if selected_item:
        item_info = master_df[master_df["ITEM"] == selected_item].iloc[0]
        st.markdown(f"""
        <div class="info-box" style="margin-bottom:0.5rem;">
            📦 <strong>{selected_item}</strong> &nbsp;|&nbsp;
            Unit: <strong>{item_info['UNIT OF MEASURE']}</strong> &nbsp;|&nbsp;
            Cost/unit: <strong>₱{float(item_info['UNIT COST']):.4f}</strong> &nbsp;|&nbsp;
            Category: <strong>{item_info['CATEGORY'].upper()}</strong>
        </div>
        """, unsafe_allow_html=True)

    if st.button("➕ Add to PO", type="secondary", use_container_width=True):
        if not selected_item:
            st.error("❌ Please select an ingredient.")
        elif quantity <= 0:
            st.error("❌ Quantity must be greater than 0.")
        else:
            st.session_state.po_items.append({
                "item": selected_item,
                "field_type": field_type,
                "quantity": quantity,
                "notes": item_notes,
                "unit": str(item_info["UNIT OF MEASURE"]),
                "category": str(item_info["CATEGORY"])
            })
            st.success(f"✅ **{selected_item}** added to PO.")
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # ── PO Cart / Review ───────────────────────────────────────────────────────
    if st.session_state.po_items:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">PO Review — {len(st.session_state.po_items)} item(s)</div>', unsafe_allow_html=True)

        # Display each item with a remove button
        for idx, po_item in enumerate(st.session_state.po_items):
            col_info, col_remove = st.columns([5, 1])
            with col_info:
                st.markdown(f"""
                <div class="po-item-row">
                    <strong>{idx+1}. {po_item['item']}</strong>
                    &nbsp;|&nbsp; <span style="color:#0f3460;">{po_item['field_type']}</span>
                    &nbsp;|&nbsp; <strong>{po_item['quantity']:,.2f} {po_item['unit']}</strong>
                    {'&nbsp;|&nbsp; 📝 ' + po_item['notes'] if po_item['notes'] else ''}
                </div>
                """, unsafe_allow_html=True)
            with col_remove:
                if st.button("🗑️", key=f"remove_{idx}", help="Remove this item"):
                    st.session_state.po_items.pop(idx)
                    st.rerun()

        st.markdown("---")

        col_clear, col_submit = st.columns([1, 2])
        with col_clear:
            if st.button("🗑️ Clear PO", use_container_width=True):
                st.session_state.po_items = []
                st.rerun()
        with col_submit:
            if st.button("💾 Submit PO", type="primary", use_container_width=True):
                if not staff_name.strip():
                    st.error("❌ Please enter your name before submitting.")
                else:
                    po_number = generate_po_number(entry_date, department)
                    log_ws = ensure_sheet(spreadsheet, LOG_SHEET, LOG_HEADERS)
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    month_lbl = entry_date.strftime("%b %Y").upper()
                    date_str = entry_date.strftime("%Y-%m-%d")

                    rows_to_log = []
                    for po_item in st.session_state.po_items:
                        ft = po_item["field_type"]
                        addin   = po_item["quantity"] if ft == "Add'l / In" else 0
                        over    = po_item["quantity"] if ft == "Over" else 0
                        damaged = po_item["quantity"] if ft == "Damaged / Out" else 0
                        cafe    = po_item["quantity"] if ft == "Café" else 0
                        bar     = po_item["quantity"] if ft == "Bar" else 0
                        kitchen = po_item["quantity"] if ft == "Kitchen Ala Carte" else 0
                        banquet = po_item["quantity"] if ft == "Kitchen Banquet" else 0
                        resto   = po_item["quantity"] if ft == "Resto" else 0

                        rows_to_log.append([
                            timestamp, month_lbl, date_str,
                            po_item["item"], staff_name.strip(),
                            po_number, department,
                            addin, over, damaged,
                            cafe, bar, kitchen, banquet, resto,
                            po_item["notes"]
                        ])

                    # Write all rows to log at once
                    log_ws.append_rows(rows_to_log)

                    # Update summary sheet for each item
                    for po_item in st.session_state.po_items:
                        ft = po_item["field_type"]
                        update_summary_for_item(
                            spreadsheet, entry_date, po_item["item"],
                            po_item["quantity"] if ft == "Add'l / In" else 0,
                            po_item["quantity"] if ft == "Over" else 0,
                            po_item["quantity"] if ft == "Damaged / Out" else 0,
                            po_item["quantity"] if ft == "Café" else 0,
                            po_item["quantity"] if ft == "Bar" else 0,
                            po_item["quantity"] if ft == "Kitchen Ala Carte" else 0,
                            po_item["quantity"] if ft == "Kitchen Banquet" else 0,
                            po_item["quantity"] if ft == "Resto" else 0,
                        )

                    item_count = len(st.session_state.po_items)
                    st.session_state.po_items = []
                    st.success(f"✅ PO **{po_number}** submitted! **{item_count} item(s)** logged by **{staff_name}** for **{department}** on **{entry_date}**.")
                    st.balloons()

        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("Your PO is empty. Search and add items above.")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: INGREDIENTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📦 Ingredients":
    st.markdown('<div class="main-header"><div><h1>📦 Ingredients</h1><p>Manage your master ingredient list</p></div></div>', unsafe_allow_html=True)

    master_df = load_master(spreadsheet)
    CATEGORIES = ["beverage","beef","chicken","pork","seafood","fresh","dry","wet","rtc","meat","frozen","dessert","other"]

    tab1, tab2, tab3 = st.tabs(["➕ Add New", "✏️ Edit / Deactivate", "📋 View All"])

    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Add New Ingredient</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            new_name = st.text_input("Item Name *")
            new_unit = st.text_input("Unit of Measure *", placeholder="gram / ml / piece / bottle / can")
        with c2:
            new_cost = st.number_input("Unit Cost (₱) *", min_value=0.0, step=0.0001, format="%.4f")
            new_cat  = st.selectbox("Category *", CATEGORIES)

        if st.button("➕ Add Ingredient", type="primary"):
            if not new_name.strip() or not new_unit.strip():
                st.error("❌ Item name and unit are required.")
            elif new_name.strip() in master_df["ITEM"].values:
                st.error("❌ This item already exists.")
            else:
                ws = ensure_sheet(spreadsheet, MASTER_SHEET, MASTER_HEADERS)
                ws.append_row([new_name.strip(), new_cost, new_unit.strip(), new_cat, "YES"])
                st.success(f"✅ **{new_name}** added to Master List!")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        if master_df.empty:
            st.info("No ingredients yet.")
        else:
            search = st.text_input("🔍 Search to Edit", placeholder="Type item name...")
            filtered_df = master_df[master_df["ITEM"].str.contains(search, case=False, na=False)] if search else master_df
            selected = st.selectbox("Select Item to Edit", filtered_df["ITEM"].tolist())

            if selected:
                item_row = master_df[master_df["ITEM"] == selected].iloc[0]
                row_idx  = master_df[master_df["ITEM"] == selected].index[0] + 2

                st.markdown('<div class="card">', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    edit_cost = st.number_input("Unit Cost (₱)", value=float(item_row["UNIT COST"]), step=0.0001, format="%.4f", key="edit_cost")
                    edit_unit = st.text_input("Unit of Measure", value=str(item_row["UNIT OF MEASURE"]), key="edit_unit")
                with c2:
                    edit_cat  = st.selectbox("Category", CATEGORIES, index=CATEGORIES.index(item_row["CATEGORY"]) if item_row["CATEGORY"] in CATEGORIES else 0, key="edit_cat")
                    edit_active = st.selectbox("Status", ["YES","NO"], index=0 if item_row["ACTIVE"]=="YES" else 1, key="edit_active")

                if st.button("💾 Save Changes", type="primary"):
                    ws = spreadsheet.worksheet(MASTER_SHEET)
                    ws.update_cell(row_idx, 2, edit_cost)
                    ws.update_cell(row_idx, 3, edit_unit)
                    ws.update_cell(row_idx, 4, edit_cat)
                    ws.update_cell(row_idx, 5, edit_active)
                    st.success(f"✅ **{selected}** updated!")
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    with tab3:
        if master_df.empty:
            st.info("No ingredients yet.")
        else:
            show_inactive = st.checkbox("Show inactive items")
            display_df = master_df if show_inactive else master_df[master_df["ACTIVE"]=="YES"]
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            st.caption(f"Total: {len(display_df)} items")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MONTH MANAGER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📅 Month Manager":
    st.markdown('<div class="main-header"><div><h1>📅 Month Manager</h1><p>Create new months and carry over ending stocks</p></div></div>', unsafe_allow_html=True)

    all_sheets = [ws.title for ws in spreadsheet.worksheets() if ws.title.startswith("SUMMARY - ")]
    months_available = [s.replace("SUMMARY - ", "") for s in all_sheets]

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Existing Months</div>', unsafe_allow_html=True)
    if months_available:
        for m in months_available:
            st.markdown(f"✅ **{m}**")
    else:
        st.info("No months set up yet. Go to Setup first.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Start a New Month</div>', unsafe_allow_html=True)
    st.write("This will create a new month summary and automatically carry over the **Ending Stocks** from the previous month as the new **Beginning Stocks**.")

    col1, col2 = st.columns(2)
    with col1:
        prev_month = st.selectbox("Previous Month (copy ending from)", months_available if months_available else ["None"])
    with col2:
        new_month_name = st.text_input("New Month Name", placeholder="e.g. JUN 2026")

    if st.button("📅 Create New Month", type="primary"):
        if not new_month_name.strip():
            st.error("❌ Please enter the new month name.")
        elif f"SUMMARY - {new_month_name.upper()}" in [ws.title for ws in spreadsheet.worksheets()]:
            st.warning(f"⚠️ {new_month_name.upper()} already exists.")
        else:
            master_df = load_master(spreadsheet)
            prev_sheet_name = f"SUMMARY - {prev_month}"
            ending_map = {}

            if prev_month != "None":
                try:
                    prev_ws = spreadsheet.worksheet(prev_sheet_name)
                    prev_records = prev_ws.get_all_records()
                    ending_map = {r["ITEM"]: float(r.get("ENDING STOCKS", 0) or 0) for r in prev_records}
                except Exception as e:
                    st.warning(f"Could not load previous month: {e}")

            new_sheet_name = f"SUMMARY - {new_month_name.upper()}"
            new_ws = spreadsheet.add_worksheet(title=new_sheet_name, rows=2000, cols=20)
            new_ws.append_row(SUMMARY_HEADERS)

            rows = []
            for _, item in master_df[master_df["ACTIVE"] == "YES"].iterrows():
                name     = item["ITEM"]
                cost     = float(item["UNIT COST"]) if item["UNIT COST"] else 0
                beginning = ending_map.get(name, 0)
                rows.append([name, cost, item["UNIT OF MEASURE"], item["CATEGORY"],
                              beginning, 0, 0, 0, 0, 0, 0, 0, 0, beginning, beginning * cost])
            if rows:
                new_ws.append_rows(rows)

            st.success(f"✅ **{new_month_name.upper()}** created with {len(rows)} items! Beginning stocks carried over from {prev_month}.")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SEARCH & HISTORY
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔍 Search & History":
    st.markdown('<div class="main-header"><div><h1>🔍 Search & History</h1><p>Find any ingredient or PO and see its full transaction log</p></div></div>', unsafe_allow_html=True)

    log_df = load_log(spreadsheet)
    master_df = load_master(spreadsheet)

    search_tab, po_tab = st.tabs(["🔍 Search by Ingredient", "📋 Search by PO Number"])

    with search_tab:
        search = st.text_input("🔍 Search Ingredient", placeholder="Type ingredient name...")

        if search:
            items_found = master_df[master_df["ITEM"].str.contains(search, case=False, na=False)]["ITEM"].tolist()
            if not items_found:
                st.warning("No ingredients found.")
            else:
                selected = st.selectbox("Select Ingredient", items_found)
                if selected:
                    item_info = master_df[master_df["ITEM"] == selected].iloc[0]

                    st.markdown(f"""
                    <div class="info-box">
                        📦 <strong>{selected}</strong> &nbsp;|&nbsp;
                        Unit: <strong>{item_info['UNIT OF MEASURE']}</strong> &nbsp;|&nbsp;
                        Cost/unit: <strong>₱{float(item_info['UNIT COST']):.4f}</strong> &nbsp;|&nbsp;
                        Category: <strong>{item_info['CATEGORY'].upper()}</strong>
                    </div>
                    """, unsafe_allow_html=True)

                    if not log_df.empty:
                        item_log = log_df[log_df["ITEM"] == selected].copy()
                        if item_log.empty:
                            st.info("No transactions recorded yet for this item.")
                        else:
                            item_log = item_log.sort_values("TIMESTAMP", ascending=False)
                            st.markdown(f"**{len(item_log)} transaction(s) found**")

                            c1, c2, c3, c4 = st.columns(4)
                            def num(col): return pd.to_numeric(item_log.get(col, 0), errors="coerce").fillna(0).sum()
                            with c1:
                                st.metric("Total Add'l/In", f"{num('ADD\'L/IN'):,.2f}")
                            with c2:
                                st.metric("Total Damaged/Out", f"{num('DAMAGED/OUT'):,.2f}")
                            with c3:
                                total_used = num("CAFÉ") + num("BAR") + num("KITCHEN ALA CARTE") + num("KITCHEN BANQUET") + num("RESTO")
                                st.metric("Total Used", f"{total_used:,.2f}")
                            with c4:
                                st.metric("Transactions", len(item_log))

                            st.markdown("### Transaction Log")
                            for _, row in item_log.iterrows():
                                movements = []
                                for col in ["ADD'L/IN","OVER","DAMAGED/OUT","CAFÉ","BAR","KITCHEN ALA CARTE","KITCHEN BANQUET","RESTO"]:
                                    val = float(row.get(col, 0) or 0)
                                    if val:
                                        movements.append(f"{col}: **{val:,.2f}**")

                                po_ref = row.get('PO NUMBER', '')
                                dept_ref = row.get('DEPARTMENT', '')
                                po_badge = f'&nbsp;|&nbsp; 🧾 <strong>{po_ref}</strong>' if po_ref else ''
                                dept_badge = f'&nbsp;|&nbsp; 🏢 {dept_ref}' if dept_ref else ''

                                st.markdown(f"""
                                <div class="log-entry">
                                    <div><strong>👤 {row.get('STAFF','—')}</strong> &nbsp;|&nbsp; 📅 {row.get('DATE','—')} &nbsp;|&nbsp; 📆 {row.get('MONTH','—')}{po_badge}{dept_badge}</div>
                                    <div style='margin-top:4px;'>{' &nbsp;·&nbsp; '.join(movements) if movements else 'No quantities entered'}</div>
                                    {'<div style="color:#888; font-size:0.8rem; margin-top:4px;">📝 ' + str(row.get('NOTES','')) + '</div>' if row.get('NOTES') else ''}
                                    <div class="timestamp">{row.get('TIMESTAMP','')}</div>
                                </div>
                                """, unsafe_allow_html=True)
                    else:
                        st.info("No transaction log yet.")

    with po_tab:
        po_search = st.text_input("🧾 Search PO Number", placeholder="e.g. PO-20260628...")

        if po_search and not log_df.empty:
            if "PO NUMBER" in log_df.columns:
                po_log = log_df[log_df["PO NUMBER"].astype(str).str.contains(po_search, case=False, na=False)]
                if po_log.empty:
                    st.warning("No PO found with that number.")
                else:
                    po_numbers = po_log["PO NUMBER"].unique()
                    for po_num in po_numbers:
                        po_entries = po_log[po_log["PO NUMBER"] == po_num]
                        first = po_entries.iloc[0]
                        st.markdown(f"""
                        <div class="info-box">
                            🧾 <strong>{po_num}</strong> &nbsp;|&nbsp;
                            📅 {first.get('DATE','—')} &nbsp;|&nbsp;
                            👤 {first.get('STAFF','—')} &nbsp;|&nbsp;
                            🏢 {first.get('DEPARTMENT','—')} &nbsp;|&nbsp;
                            <strong>{len(po_entries)} item(s)</strong>
                        </div>
                        """, unsafe_allow_html=True)

                        for _, row in po_entries.iterrows():
                            movements = []
                            for col in ["ADD'L/IN","OVER","DAMAGED/OUT","CAFÉ","BAR","KITCHEN ALA CARTE","KITCHEN BANQUET","RESTO"]:
                                val = float(row.get(col, 0) or 0)
                                if val:
                                    movements.append(f"{col}: **{val:,.2f}**")
                            st.markdown(f"""
                            <div class="log-entry">
                                <div><strong>📦 {row.get('ITEM','—')}</strong></div>
                                <div style='margin-top:4px;'>{' &nbsp;·&nbsp; '.join(movements) if movements else 'No quantities'}</div>
                                {'<div style="color:#888; font-size:0.8rem; margin-top:4px;">📝 ' + str(row.get('NOTES','')) + '</div>' if row.get('NOTES') else ''}
                            </div>
                            """, unsafe_allow_html=True)
            else:
                st.info("No PO data found. Old entries may not have PO numbers.")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📊 Summary":
    st.markdown('<div class="main-header"><div><h1>📊 Summary</h1><p>Live inventory summary by month</p></div></div>', unsafe_allow_html=True)

    all_sheets = [ws.title for ws in spreadsheet.worksheets() if ws.title.startswith("SUMMARY - ")]
    months_available = [s.replace("SUMMARY - ", "") for s in all_sheets]

    if not months_available:
        st.warning("No months set up yet. Please go to Setup first.")
    else:
        selected_month = st.selectbox("Select Month", months_available)
        cat_filter = st.multiselect("Filter by Category", ["All","beverage","beef","chicken","pork","seafood","fresh","dry","wet","rtc","meat","frozen","dessert"], default=["All"])

        sum_ws = spreadsheet.worksheet(f"SUMMARY - {selected_month}")
        records = sum_ws.get_all_records()
        df = pd.DataFrame(records) if records else pd.DataFrame(columns=SUMMARY_HEADERS)

        if not df.empty:
            if "All" not in cat_filter and cat_filter:
                df = df[df["CATEGORY"].isin(cat_filter)]

            total_worth = pd.to_numeric(df["TOTAL WORTH OF STOCKS"], errors="coerce").fillna(0).sum()
            total_items = len(df[pd.to_numeric(df["ENDING STOCKS"], errors="coerce").fillna(0) > 0])

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Total Items with Stock", total_items)
            with c2:
                st.metric("Total Worth of Stocks", f"₱{total_worth:,.2f}")
            with c3:
                st.metric("Month", selected_month)

            st.markdown("---")
            numeric_cols = ["BEGINNING STOCKS","ADD'L/IN","OVER","DAMAGED/OUT","CAFÉ","BAR","KITCHEN ALA CARTE","KITCHEN BANQUET","RESTO","ENDING STOCKS","TOTAL WORTH OF STOCKS"]
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            st.dataframe(df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: EXPORT TO EXCEL
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⬇️ Export to Excel":
    st.markdown('<div class="main-header"><div><h1>⬇️ Export to Excel</h1><p>Download monthly inventory in your original format</p></div></div>', unsafe_allow_html=True)

    all_sheets = [ws.title for ws in spreadsheet.worksheets() if ws.title.startswith("SUMMARY - ")]
    months_available = [s.replace("SUMMARY - ", "") for s in all_sheets]

    if not months_available:
        st.warning("No months available to export.")
    else:
        selected_month = st.selectbox("Select Month to Export", months_available)

        if st.button("📥 Generate Excel File", type="primary"):
            sum_ws = spreadsheet.worksheet(f"SUMMARY - {selected_month}")
            records = sum_ws.get_all_records()
            df = pd.DataFrame(records) if records else pd.DataFrame(columns=SUMMARY_HEADERS)

            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "SUMMARY"

            header_fill = PatternFill("solid", fgColor="1a1a2e")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin = Side(style="thin", color="CCCCCC")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            ws_out.merge_cells("A1:O1")
            title_cell = ws_out["A1"]
            title_cell.value = f"SERVANDO MAIN WAREHOUSE INVENTORY — {selected_month}"
            title_cell.font = Font(bold=True, size=14, color="FFFFFF")
            title_cell.fill = PatternFill("solid", fgColor="0f3460")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_out.row_dimensions[1].height = 30

            for col_idx, header in enumerate(SUMMARY_HEADERS, 1):
                cell = ws_out.cell(row=2, column=col_idx, value=header)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border
            ws_out.row_dimensions[2].height = 40

            cat_colors = {
                "beverage": "EBF5FB", "beef": "FDEDEC", "chicken": "FEF9E7",
                "seafood": "E8F8F5", "fresh": "E9F7EF", "dry": "FDF2E9",
                "wet": "F4ECF7", "rtc": "EAF2FF", "pork": "FDEDEC",
                "meat": "FDFEFE", "frozen": "EBF5FB", "dessert": "FEF9E7"
            }
            num_fmt = '#,##0.0000'

            for row_idx, (_, row) in enumerate(df.iterrows(), 3):
                cat = str(row.get("CATEGORY", "")).lower()
                fill_color = cat_colors.get(cat, "FFFFFF")
                row_fill = PatternFill("solid", fgColor=fill_color)

                for col_idx, header in enumerate(SUMMARY_HEADERS, 1):
                    val = row.get(header, "")
                    cell = ws_out.cell(row=row_idx, column=col_idx, value=val)
                    cell.fill = row_fill
                    cell.border = border
                    if col_idx in [2, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]:
                        cell.number_format = num_fmt
                        cell.alignment = Alignment(horizontal="right")
                    else:
                        cell.alignment = Alignment(vertical="center")

            total_row = len(df) + 3
            ws_out.cell(row=total_row, column=1, value="TOTAL WORTH OF STOCKS").font = Font(bold=True)
            ws_out.cell(row=total_row, column=1).fill = PatternFill("solid", fgColor="1a1a2e")
            ws_out.cell(row=total_row, column=1).font = Font(bold=True, color="FFFFFF")

            total_worth = pd.to_numeric(df["TOTAL WORTH OF STOCKS"], errors="coerce").fillna(0).sum()
            tw_cell = ws_out.cell(row=total_row, column=15, value=round(total_worth, 4))
            tw_cell.fill = PatternFill("solid", fgColor="0f3460")
            tw_cell.font = Font(bold=True, color="FFFFFF")
            tw_cell.number_format = '#,##0.0000'

            col_widths = [35, 12, 16, 12, 16, 10, 10, 14, 10, 10, 18, 18, 10, 15, 20]
            for i, w in enumerate(col_widths, 1):
                ws_out.column_dimensions[get_column_letter(i)].width = w

            ws_out.freeze_panes = "A3"

            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)

            st.download_button(
                label=f"📥 Download {selected_month} Inventory.xlsx",
                data=buf,
                file_name=f"SERVANDO_INVENTORY_{selected_month.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ Excel file ready for download!")