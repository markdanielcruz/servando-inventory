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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a1a2e 0%, #16213e 60%, #0f3460 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}
section[data-testid="stSidebar"] * { color: #e0e0e0 !important; }
section[data-testid="stSidebar"] .stRadio label { 
    font-size: 0.95rem; 
    padding: 6px 0;
}

/* Header */
.main-header {
    background: linear-gradient(135deg, #1a1a2e 0%, #0f3460 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.main-header h1 {
    font-family: 'Playfair Display', serif;
    font-size: 1.8rem;
    margin: 0;
    color: white;
}
.main-header p { margin: 0; color: rgba(255,255,255,0.7); font-size: 0.9rem; }

/* Cards */
.card {
    background: white;
    border-radius: 12px;
    padding: 1.5rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    border: 1px solid #f0f0f0;
    margin-bottom: 1rem;
}
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    text-align: center;
}
.metric-card .value { font-size: 2rem; font-weight: 700; }
.metric-card .label { font-size: 0.8rem; opacity: 0.85; margin-top: 4px; }

/* Buttons */
.stButton > button {
    border-radius: 8px;
    font-weight: 500;
    transition: all 0.2s;
}
.stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }

/* Input fields */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stSelectbox > div > div {
    border-radius: 8px;
    border: 1.5px solid #e0e0e0;
}

/* Section titles */
.section-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #1a1a2e;
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #0f3460;
    display: inline-block;
}

/* Log entry */
.log-entry {
    background: #f8f9ff;
    border-left: 3px solid #0f3460;
    padding: 0.75rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 0.5rem;
    font-size: 0.9rem;
}
.log-entry .timestamp { color: #888; font-size: 0.8rem; }

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
.badge-beverage { background: #e8f4fd; color: #1a6fa8; }
.badge-dry { background: #fef9e7; color: #b7950b; }
.badge-fresh { background: #e9f7ef; color: #1e8449; }
.badge-wet { background: #e8f8f5; color: #148f77; }
.badge-beef { background: #fde8e8; color: #a93226; }
.badge-chicken { background: #fef5e7; color: #ca6f1e; }
.badge-seafood { background: #e8f4fd; color: #1f618d; }
.badge-pork { background: #fdedec; color: #c0392b; }
.badge-rtc { background: #f4ecf7; color: #7d3c98; }

/* Success/info boxes */
.success-box {
    background: #e9f7ef;
    border: 1px solid #a9dfbf;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    color: #1e8449;
    margin: 0.5rem 0;
}
.info-box {
    background: #eaf2ff;
    border: 1px solid #aed6f1;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    color: #1a5276;
    margin: 0.5rem 0;
}

/* Table styling */
.dataframe { border-radius: 8px; overflow: hidden; }
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
LOG_HEADERS     = ["TIMESTAMP", "MONTH", "DATE", "ITEM", "STAFF",
                   "ADD'L/IN", "OVER", "DAMAGED/OUT", "CAFÉ", "BAR",
                   "KITCHEN ALA CARTE", "KITCHEN BANQUET", "RESTO", "NOTES"]
SUMMARY_HEADERS = ["ITEM", "UNIT COST", "UNIT OF MEASURE", "CATEGORY",
                   "BEGINNING STOCKS", "ADD'L/IN", "OVER", "DAMAGED/OUT",
                   "CAFÉ", "BAR", "KITCHEN ALA CARTE", "KITCHEN BANQUET",
                   "RESTO", "ENDING STOCKS", "TOTAL WORTH OF STOCKS"]

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

        # Get beginning stocks from summary sheet meta row (stored separately)
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
    <div style='text-align:center; padding: 1rem 0 1.5rem 0;'>
        <div style='font-size:2.5rem;'>🏪</div>
        <div style='font-family: Playfair Display, serif; font-size:1.2rem; font-weight:700; color:white; margin-top:0.5rem;'>Servando</div>
        <div style='font-size:0.75rem; color:rgba(255,255,255,0.5); letter-spacing:2px; text-transform:uppercase;'>Inventory System</div>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("Navigation", [
        "📋 Daily Input",
        "📦 Ingredients",
        "📅 Month Manager",
        "🔍 Search & History",
        "📊 Summary",
        "⬇️ Export to Excel",
        "⚙️ Setup"
    ])

    st.markdown("---")
    st.markdown(f"<div style='font-size:0.75rem; color:rgba(255,255,255,0.4);'>Today: {date.today().strftime('%B %d, %Y')}</div>", unsafe_allow_html=True)

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
# PAGE: DAILY INPUT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Daily Input":
    st.markdown('<div class="main-header"><div><h1>📋 Daily Input</h1><p>Record inventory movements for today</p></div></div>', unsafe_allow_html=True)

    master_df = load_master(spreadsheet)
    if master_df.empty:
        st.warning("⚠️ No ingredients found. Please go to Setup first.")
        st.stop()

    active_items = master_df[master_df["ACTIVE"] == "YES"]["ITEM"].tolist()

    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        entry_date = st.date_input("📅 Date", value=date.today())
    with col2:
        staff_name = st.text_input("👤 Staff Name", placeholder="Enter your name")
    with col3:
        selected_month = st.text_input("📆 Month (auto-filled)", value=entry_date.strftime("%b %Y").upper(), disabled=True)

    st.markdown("---")

    search_query = st.text_input("🔍 Search Ingredient", placeholder="Type to search... e.g. 'chicken', 'beef'")
    filtered = [i for i in active_items if search_query.lower() in i.lower()] if search_query else active_items

    if search_query and not filtered:
        st.warning("No ingredients found matching your search.")
    else:
        if search_query:
            st.caption(f"Found {len(filtered)} item(s)")

        selected_item = st.selectbox("Select Ingredient", filtered)

        if selected_item:
            item_info = master_df[master_df["ITEM"] == selected_item].iloc[0]
            st.markdown(f"""
            <div class="info-box">
                📦 <strong>{selected_item}</strong> &nbsp;|&nbsp; 
                Unit: <strong>{item_info['UNIT OF MEASURE']}</strong> &nbsp;|&nbsp;
                Cost/unit: <strong>₱{float(item_info['UNIT COST']):.4f}</strong> &nbsp;|&nbsp;
                Category: <strong>{item_info['CATEGORY'].upper()}</strong>
            </div>
            """, unsafe_allow_html=True)

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Enter Quantities</div>', unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)
            with c1:
                addin    = st.number_input("➕ Add'l / In",      min_value=0.0, step=0.01, format="%.2f")
                over     = st.number_input("🔼 Over",             min_value=0.0, step=0.01, format="%.2f")
                damaged  = st.number_input("❌ Damaged / Out",    min_value=0.0, step=0.01, format="%.2f")
            with c2:
                cafe     = st.number_input("☕ Café",             min_value=0.0, step=0.01, format="%.2f")
                bar      = st.number_input("🍸 Bar",              min_value=0.0, step=0.01, format="%.2f")
                resto    = st.number_input("🍽️ Resto",            min_value=0.0, step=0.01, format="%.2f")
            with c3:
                kitchen  = st.number_input("🍳 Kitchen Ala Carte",min_value=0.0, step=0.01, format="%.2f")
                banquet  = st.number_input("🎉 Kitchen Banquet",  min_value=0.0, step=0.01, format="%.2f")
                notes    = st.text_input("📝 Notes (optional)")

            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("💾 Save Entry", type="primary", use_container_width=True):
                if not staff_name.strip():
                    st.error("❌ Please enter your name before saving.")
                else:
                    log_ws = ensure_sheet(spreadsheet, LOG_SHEET, LOG_HEADERS)
                    row = [
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        entry_date.strftime("%b %Y").upper(),
                        entry_date.strftime("%Y-%m-%d"),
                        selected_item,
                        staff_name.strip(),
                        addin, over, damaged,
                        cafe, bar, kitchen, banquet, resto,
                        notes
                    ]
                    log_ws.append_row(row)

                    # Update summary sheet
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

                    st.success(f"✅ Entry saved! **{selected_item}** logged by **{staff_name}** on **{entry_date}**")
                    st.balloons()

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

    # List existing summary sheets
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
    st.markdown('<div class="main-header"><div><h1>🔍 Search & History</h1><p>Find any ingredient and see its full transaction log</p></div></div>', unsafe_allow_html=True)

    log_df = load_log(spreadsheet)
    master_df = load_master(spreadsheet)

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

                        # Summary metrics
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

                            st.markdown(f"""
                            <div class="log-entry">
                                <div><strong>👤 {row.get('STAFF','—')}</strong> &nbsp;|&nbsp; 📅 {row.get('DATE','—')} &nbsp;|&nbsp; 📆 {row.get('MONTH','—')}</div>
                                <div style='margin-top:4px;'>{' &nbsp;·&nbsp; '.join(movements) if movements else 'No quantities entered'}</div>
                                {'<div style="color:#888; font-size:0.8rem; margin-top:4px;">📝 ' + str(row.get('NOTES','')) + '</div>' if row.get('NOTES') else ''}
                                <div class="timestamp">{row.get('TIMESTAMP','')}</div>
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.info("No transaction log yet.")

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

            # Metrics row
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

            # Header styling
            header_fill = PatternFill("solid", fgColor="1a1a2e")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin = Side(style="thin", color="CCCCCC")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            # Title row
            ws_out.merge_cells("A1:O1")
            title_cell = ws_out["A1"]
            title_cell.value = f"SERVANDO MAIN WAREHOUSE INVENTORY — {selected_month}"
            title_cell.font = Font(bold=True, size=14, color="FFFFFF")
            title_cell.fill = PatternFill("solid", fgColor="0f3460")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_out.row_dimensions[1].height = 30

            # Headers row
            for col_idx, header in enumerate(SUMMARY_HEADERS, 1):
                cell = ws_out.cell(row=2, column=col_idx, value=header)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border
            ws_out.row_dimensions[2].height = 40

            # Data rows
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

            # Total row
            total_row = len(df) + 3
            ws_out.cell(row=total_row, column=1, value="TOTAL WORTH OF STOCKS").font = Font(bold=True)
            ws_out.cell(row=total_row, column=1).fill = PatternFill("solid", fgColor="1a1a2e")
            ws_out.cell(row=total_row, column=1).font = Font(bold=True, color="FFFFFF")

            total_worth = pd.to_numeric(df["TOTAL WORTH OF STOCKS"], errors="coerce").fillna(0).sum()
            tw_cell = ws_out.cell(row=total_row, column=15, value=round(total_worth, 4))
            tw_cell.fill = PatternFill("solid", fgColor="0f3460")
            tw_cell.font = Font(bold=True, color="FFFFFF")
            tw_cell.number_format = '#,##0.0000'

            # Column widths
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
