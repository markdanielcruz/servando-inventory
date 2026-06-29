import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Servando Warehouse Inventory",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ────────────────────────────────────────────────────────────────────────
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
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='900' height='700' viewBox='0 0 900 700'%3E%3Crect width='900' height='700' fill='%230A1208'/%3E%3Cg opacity='0.28'%3E%3Cpath d='M-10,700 C30,580 110,510 165,415 C192,362 180,298 155,258 C198,302 214,370 192,425 C242,358 252,265 230,196 C272,275 260,378 228,445 C276,368 298,254 286,162 C324,250 300,376 262,458 C235,512 196,596 168,700Z' fill='%231A4A1F'/%3E%3C/g%3E%3Cg opacity='0.22' transform='translate(860,0)'%3E%3Cpath d='M20,700 C0,580 -42,498 -78,415 C-100,358 -88,288 -64,242 C-98,292 -110,368 -92,425 C-130,352 -136,254 -112,185 C-152,270 -140,386 -106,450 C-140,362 -172,244 -166,138 C-196,242 -168,382 -128,462 C-104,518 -70,600 -46,700Z' fill='%231A4A1F'/%3E%3C/g%3E%3Cg opacity='0.18' transform='translate(370,-70)'%3E%3Cpath d='M70,0 C26,92 -4,206 12,322 C-28,258 -40,162 -12,80 C-62,126 -74,242 -46,334 C-86,270 -82,160 -56,80 C-104,160 -92,292 -56,390 C-28,430 18,458 64,452 C108,446 150,406 164,350 C182,270 166,154 124,80 C144,160 138,276 112,338 C130,248 126,132 98,58 C78,-16 74,-16 70,0Z' fill='%23163D1A'/%3E%3C/g%3E%3C/svg%3E");
    background-attachment: fixed;
    background-size: cover;
}
[data-testid="stSidebar"] {
    background-color: rgba(8,14,8,0.95) !important;
    border-right: 1px solid #2A3828;
    backdrop-filter: blur(8px);
}
[data-testid="stSidebar"] * { color: #D4D0C8 !important; }
.main-header {
    background: linear-gradient(135deg, rgba(26,46,26,0.85) 0%, rgba(21,37,21,0.85) 100%);
    border: 1px solid #2A3828;
    border-radius: 12px;
    padding: 20px 28px;
    margin-bottom: 20px;
    backdrop-filter: blur(6px);
}
.main-header h1 {
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.7rem;
    color: #8CAF7A;
    margin: 0 0 2px 0;
    font-weight: 600;
}
.main-header p { font-size: 0.7rem; letter-spacing: 2.5px; text-transform: uppercase; color: #4A6B3E; margin: 0; }
.card {
    background: #111E11;
    border: 1px solid #1E2E1C;
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}
.section-title {
    font-size: 0.62rem;
    font-weight: 600;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #5A7A52;
    margin-bottom: 0.75rem;
    padding-bottom: 6px;
    border-bottom: 1px solid #1E2E1C;
    display: block;
}
.stButton > button {
    background-color: #2A3E28 !important;
    color: #A8C896 !important;
    border: 1px solid #3A5238 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background-color: #3A5238 !important;
    border-color: #5A7A52 !important;
    color: #C8DCC0 !important;
}
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
label { color: #8A9E84 !important; font-size: 0.78rem !important; font-weight: 500 !important; }
.log-entry {
    background: #0E1810;
    border-left: 3px solid #3A5238;
    padding: 0.6rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 0.4rem;
    font-size: 0.88rem;
    color: #D4D0C8;
}
.po-item-row {
    background: #0E1810;
    border: 1px solid #2A3828;
    border-radius: 8px;
    padding: 0.6rem 1rem;
    margin-bottom: 0.4rem;
    color: #D4D0C8;
}
.info-box {
    background: #111E11;
    border: 1px solid #2A3828;
    border-radius: 8px;
    padding: 0.6rem 1rem;
    color: #8CAF7A;
    margin: 0.4rem 0;
    font-size: 0.88rem;
}
.success-box {
    background: #1A2E1A;
    border: 1px solid #3A5238;
    border-radius: 8px;
    padding: 0.6rem 1rem;
    color: #A8C896;
    margin: 0.4rem 0;
}
[data-testid="metric-container"] {
    background: #111E11 !important;
    border: 1px solid #1E2E1C !important;
    border-radius: 10px !important;
    padding: 12px 14px !important;
}
[data-testid="metric-container"] label {
    font-size: 0.6rem !important; letter-spacing: 2px !important;
    text-transform: uppercase !important; color: #5A7A52 !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Cormorant Garamond', serif !important;
    font-size: 1.5rem !important; font-weight: 600 !important; color: #A8C896 !important;
}
[data-testid="stDataFrame"] { border: 1px solid #1E2E1C !important; border-radius: 10px !important; }
hr { border-color: #1E2E1C !important; }
h3 { font-family: 'Cormorant Garamond', serif !important; color: #A8C896 !important; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1pYqecDqe_qUwkexRiG5mmOaYV58_1uoRCx_Yy4CepNQ"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

ITEMS_SHEET    = "v2_ITEMS"
LOG_SHEET      = "v2_LOG"

ITEMS_HEADERS  = ["ITEM", "UNIT COST", "UNIT OF MEASURE", "CATEGORY", "BEGINNING_STOCKS", "ACTIVE"]
LOG_HEADERS    = ["TIMESTAMP", "DATE", "MONTH", "ITEM", "STAFF", "TXN_TYPE", "REF_NUMBER",
                  "ADD_IN", "OVER", "RESTAURANT", "BANQUET", "CAFE", "BAR", "OTHERS", "SPOILAGE", "NOTES"]

DEPARTMENTS    = ["Restaurant", "Banquet", "Café", "Bar", "Others"]
CATEGORIES     = ["beverage","beef","chicken","pork","seafood","fresh","dry","wet","rtc","meat","frozen","dessert","other"]

# ── 557 items from May 2026 ────────────────────────────────────────────────────
INITIAL_ITEMS = [
('Absolut Vodka Blue', 5100.0, 0.92, 'ml', 'beverage'),
('Accord Powder', 0.0, 0.19, 'gram', 'dry'),
('Agave Syrup 1.02kg', 14280.0, 1.088235294117647, 'ml', 'wet'),
('Almond - Ground', 3000.0, 1.08, 'gram', 'dry'),
('Amaretto Disaronno 700 ml', 5600.0, 1.32, 'ml', 'beverage'),
('Ampalaya - Fresh', 0.0, 0.18, 'gram', 'fresh'),
('Ampalaya - Dahon', 0.0, 0.41, 'gram', 'fresh'),
('Anchovies', 448.0, 3.3035714285714284, 'gram', 'wet'),
('Anisado Wine', 0.0, 0.21428571428571427, 'ml', 'wet'),
('Aperol 700 ml', 4900.0, 1.2857142857142858, 'ml', 'beverage'),
('Apog Powder', 0.0, 3.01, 'gram', 'dry'),
('Apple - Red', 0.0, 0.16, 'gram', 'fresh'),
('Atsuete / Annatto -  Seeds', 0.0, 0.85, 'gram', 'dry'),
('Atsuete / Annatto - Powdered', 0.0, 0.33, 'gram', 'dry'),
('Bacardi - Gold Rum 750ml', 1500.0, 0.9333333333333333, 'ml', 'beverage'),
('Bacardi - Premium Black 750ml', 3000.0, 1.1666666666666667, 'ml', 'beverage'),
('Bacardi - Superior White 750 ml', 1500.0, 2.3066666666666666, 'ml', 'beverage'),
("Bailey's Original 700 ml", 3500.0, 1.1428571428571428, 'bottle', 'beverage'),
("Baker's Best Margarine 225g", 4050.0, 0.28, 'gram', 'wet'),
('Baking Powder', 1000.0, 0.22, 'gram', 'dry'),
('Baking Soda', 4535.0, 0.4852941176470588, 'gram', 'dry'),
('Balls - Mozarella', 0.0, 1.16, 'gram', 'rtc'),
('Balls - Shrimp', 0.0, 0.215, 'gram', 'rtc'),
('Balls - Squid', 0.0, 0.297, 'gram', 'rtc'),
('Bambi Spring Roll Wrapper - Big 25 pcs', 0.0, 10.0, 'piece', 'rtc'),
('Bambi Spring Roll Wrapper Small - 35 pcs', 0.0, 8.0, 'piece', 'rtc'),
('Banana - Lakatan', 0.0, 0.14, 'gram', 'fresh'),
('Banana - Saba', 0.0, 0.08, 'gram', 'fresh'),
('Banana - Sweetened, for Halo halo per gram', 13000.0, 0.19, 'gram', 'fresh'),
('Banana Blossom', 0.0, 0.57, 'gram', 'dry'),
('Banana Essence', 0.0, 3.75, 'ml', 'wet'),
('Banana Ketchup', 0.0, 0.77, 'ml', 'wet'),
('Bangus - Daing (pack of 3, per piece)', 0.0, 75.0, 'piece', 'seafood'),
('Bangus - Smoked Milkfish, Tinapa (420 grams)', 9.0, 360.0, 'piece', 'seafood'),
('Bangus - Unseasoned Fillet - 520 grams', 13.0, 298.0, 'piece', 'seafood'),
('Barbeque Sauce - Ready made', 0.0, 2.2, 'ml', 'wet'),
('Barrio Fiesta Bagoong Regular', 7000.0, 0.45, 'gram', 'wet'),
('Basil Leaves - Dried', 0.0, 1.49, 'gram', 'dry'),
('Bay Leaves', 0.0, 3.0, 'gram', 'dry'),
('Beans - for Halo halo per gram', 12000.0, 0.24, 'gram', 'wet'),
('Beef - Brisket', 78670.0, 0.48, 'gram', 'beef'),
('Beef - Feet', 0.0, 0.24, 'gram', 'beef'),
('Beef - Ground', 5000.0, 0.395, 'gram', 'beef'),
('Beef - Shank', 40120.09, 0.43, 'gram', 'beef'),
('Beef - Skin', 5000.0, 0.22, 'gram', 'beef'),
('Belgian Waffle Mix', 0.0, 0.328, 'gram', 'dry'),
('Bell Pepper - Green', 0.0, 0.3, 'gram', 'fresh'),
('Bell Pepper - Red', 0.0, 0.3, 'gram', 'fresh'),
('Black Beans', 0.0, 0.47, 'gram', 'wet'),
('Black Olives - Pitted, Sliced', 6075.0, 0.6977777777777778, 'gram', 'wet'),
('Black Pepper - Ground', 1000.0, 0.698, 'gram', 'dry'),
('Black Pepper - Whole', 0.0, 1.2, 'gram', 'dry'),
('Blue Curacao 1 liter', 2000.0, 0.52, 'ml', 'beverage'),
('Blue Lemonade Powdered Juice 500g', 2000.0, 1.175, 'gram', 'beverage'),
('Bombay Sapphire Gin 750 ml', 1900.0, 1.6, 'ml', 'beverage'),
('Bread Crumbs - Fine', 0.0, 0.18, 'gram', 'dry'),
('Bread Crumbs - Japanese', 3000.0, 0.18, 'gram', 'dry'),
('Broccoli', 0.0, 0.12, 'gram', 'fresh'),
('Broth Cubes - Beef', 2220.0, 0.75, 'gram', 'rtc'),
('Broth Cubes - Chicken', 600.0, 0.75, 'gram', 'rtc'),
('Broth Cubes - Pork', 840.0, 0.75, 'gram', 'rtc'),
('Broth Cubes - Shrimp', 2400.0, 0.75, 'gram', 'rtc'),
('Butter - Unsalted', 0.0, 0.3, 'gram', 'rtc'),
('Cabbage', 0.0, 0.1, 'gram', 'fresh'),
('Cabbage - Red', 0.0, 0.42, 'gram', 'fresh'),
('Cajun Powder', 0.0, 1.7875, 'gram', 'wet'),
('Calamansi - Concentrate, with Honey', 0.0, 0.48, 'ml', 'wet'),
('Calamansi - Fresh', 0.0, 0.14, 'gram', 'fresh'),
('Calamansi - Pure, Unsweetened', 1000.0, 0.2, 'ml', 'wet'),
('Campari Bitter 750 ml', 3000.0, 1.56, 'ml', 'beverage'),
('Canada Dry Ginger Ale 355 ml', 36.0, 50.0, 'can', 'beverage'),
('Canned Tuna', 0.0, 0.361, 'can', 'rtc'),
('Capers', 400.0, 0.9614035087719298, 'gram', 'wet'),
('Caramel Sauce - Torani', 0.0, 0.65, 'ml', 'wet'),
('Carrots', 0.0, 0.19, 'gram', 'fresh'),
('Casa Okinawa Flavor', 0.0, 0.43, 'gram', 'dry'),
('Cascina Italian Pesto 540 g per bottle', 9.0, 394.0, 'jar', 'wet'),
('Cauliflower', 0.0, 0.32, 'gram', 'fresh'),
('Cayenne pepper', 450.0, 2.8846153846153846, 'gram', 'dry'),
('Celery', 0.0, 0.18, 'gram', 'fresh'),
('Cheese  - Sauce (500g)', 0.0, 0.65, 'ml', 'wet'),
('Cheese - Cheddar - Block', 10020.0, 0.4025, 'gram', 'rtc'),
('Cheese - Filled', 2700.0, 0.31, 'gram', 'rtc'),
('Cheese - Mozarella - Block', 38890.0, 0.5, 'gram', 'rtc'),
('Cheese - Parmesan - Ground', 10938.0, 1.18, 'gram', 'rtc'),
('Cheese - Parmesan Block', 0.0, 1.66, 'gram', 'rtc'),
('Cheese - Quickmelt', 3600.0, 0.4633333333333333, 'gram', 'rtc'),
('Chicken - Breast Fillet', 12000.0, 0.26, 'gram', 'chicken'),
('Chicken - Breast, Bone In', 0.0, 0.24, 'gram', 'chicken'),
('Chicken - Leg Quarter Bone In', 15000.0, 0.21, 'gram', 'chicken'),
('Chicken - Leg Quarter Fillet Boneless', 56000.0, 0.335, 'gram', 'chicken'),
('Chicken - Liver', 6000.0, 0.265, 'gram', 'chicken'),
('Chicken - Skin', 0.0, 0.175, 'gram', 'chicken'),
('Chicken - Whole', 0.0, 0.235, 'gram', 'chicken'),
('Chicken - Wings', 0.0, 0.18, 'gram', 'chicken'),
('Chili - Flakes', 0.0, 0.55, 'gram', 'dry'),
('Chili - Green', 0.0, 0.25, 'gram', 'fresh'),
('Chili - Labuyo', 0.0, 0.2, 'gram', 'fresh'),
('Chili - Leaves', 0.0, 0.12, 'gram', 'dry'),
('Chili - Powder', 0.0, 0.27, 'gram', 'dry'),
('Chorizo - Bilbao', 0.0, 1.9857142857142858, 'gram', 'meat'),
('Chorizo - Macao', 0.0, 0.6363636363636364, 'gram', 'meat'),
('Chorizo - Spanish', 0.0, 0.65, 'gram', 'meat'),
('Cinnamon - Ground', 200.0, 0.745, 'gram', 'dry'),
('Cinnamon Sticks', 539.0, 1.39, 'gram', 'dry'),
('Cloves - Powder', 0.0, 2.744186046511628, 'gram', 'dry'),
('Cocoa Powder', 1000.0, 2.331288343558282, 'gram', 'dry'),
('Coconut - Shredded', 0.0, 55.0, 'gram', 'fresh'),
('Coconut Cream per can', 8.0, 55.0, 'can', 'wet'),
('Coconut Rum - Malibu', 5600.0, 1.0271428571428571, 'ml', 'beverage'),
('Coconut Water - Vita Coco', 0.0, 0.15151515151515152, 'ml', 'beverage'),
('Coffee Beans - Silcafe', 0.0, 0.73, 'gram', 'dry'),
('Coke 1.5L per ml', 0.0, 0.095, 'ml', 'beverage'),
('Coke Original 320 ml', 50.0, 37.0, 'can', 'beverage'),
('Coke Zero 320 ml', 67.0, 37.0, 'can', 'beverage'),
('Cooking Sake1 Liter per ml', 0.0, 0.23, 'ml', 'wet'),
('Coriander Powder', 0.0, 0.49, 'gram', 'dry'),
('Corn - Kernels, Canned per gram', 425.0, 0.14285714285714285, 'gram', 'wet'),
('Corn - Sweet Style, Canned per gram', 425.0, 0.18823529411764706, 'gram', 'wet'),
('Corn Flakes', 0.0, 0.3073770491803279, 'gram', 'dry'),
('Cornstarch', 0.0, 0.07, 'gram', 'dry'),
('Corona Beer 330 ml', 144.0, 70.0, 'bottle', 'beverage'),
('Crab Meat, Frozen', 5000.0, 1.327, 'gram', 'seafood'),
('Crab Paste', 908.0, 0.8810572687224669, 'gram', 'wet'),
('Crab Stick', 0.0, 0.329, 'gram', 'seafood'),
('Cream - All Purpose, 250ml per ml', 12000.0, 0.32, 'ml', 'wet'),
('Cream - Coconut per ml', 57200.0, 0.2125, 'ml', 'wet'),
('Cream - Cooking', 0.0, 0.35, 'ml', 'wet'),
('Cream - Full Whipping', 0.0, 0.12, 'liter', 'wet'),
('Cream - Mascarpone', 11000.0, 0.89, 'ml', 'wet'),
('Cream - Sour', 0.0, 0.496, 'ml', 'wet'),
('Cream Cheese', 1125.0, 0.8666666666666667, 'gram', 'rtc'),
('Cream Dory - Fillet', 3000.0, 0.135, 'gram', 'seafood'),
('Cream Of Mushroom, Powdered', 1550.0, 0.8064516129032258, 'gram', 'wet'),
('Cream of Tartar', 0.0, 0.962, 'gram', 'dry'),
('Creamer Sachet (100 pcs, per piece)', 500.0, 2.0, 'piece', 'dry'),
('Crushed Graham', 1000.0, 0.225, 'gram', 'dry'),
('Cucumber - Fresh', 0.0, 0.13, 'gram', 'fresh'),
('Cumin - Ground', 600.0, 0.842, 'gram', 'dry'),
('Curacao Triple Sec 750 ml', 0.0, 0.6, 'ml', 'beverage'),
('Custard Powder - Cremyvit', 18000.0, 0.51, 'gram', 'dry'),
('Dari Cream Margarine 200g', 0.0, 0.47, 'gram', 'rtc'),
('Dark Chocolate - Bar', 1000.0, 0.389, 'gram', 'dry'),
('Dark Chocolate - Droplets', 1000.0, 0.558, 'gram', 'dry'),
('Dark Chocolate Sauce - Torani', 0.0, 0.7, 'ml', 'wet'),
('Dessicated Coconut', 0.0, 0.325, 'gram', 'dry'),
('Dijon Mustard', 1362.0, 0.914, 'bottle', 'wet'),
('Dill - Fresh', 0.0, 0.81, 'gram', 'fresh'),
('Double Chocolate Powder', 0.0, 1.01, 'gram', 'dry'),
('Easy Cream Base (Frappe Base) per gram', 10000.0, 0.405, 'gram', 'dry'),
('Eden Cheese per gram', 0.0, 0.372, 'gram', 'rtc'),
('Edible Leaf Foil Gold Flakes per gram', 36.0, 16.0, 'gram', 'dry'),
('Egg - Fresh (per piece)', 0.0, 12.0, 'piece', 'fresh'),
('Egg - Salted (per piece)', 0.0, 20.0, 'piece', 'fresh'),
('Eggplant - Fresh', 0.0, 0.13, 'gram', 'fresh'),
('Elderflower Liqueur', 0.0, 2.7857142857142856, 'ml', 'beverage'),
('Evian Water', 0.0, 95.0, 'bottle', 'beverage'),
('Feuilletine', 0.0, 2.67, 'gram', 'dry'),
('Fish Sauce (Patis)', 6000.0, 0.08, 'ml', 'wet'),
('Five Spice Powder', 0.0, 0.425, 'gram', 'dry'),
('Flour - All Purpose', 25000.0, 0.0584, 'gram', 'dry'),
('Flour - Bread', 0.0, 0.065, 'gram', 'dry'),
('Flour - Cake', 2000.0, 0.07, 'gram', 'dry'),
('Flour - Caputo (00 Flour)', 25000.0, 0.16, 'gram', 'dry'),
('Flour - Glutinous Rice', 0.0, 0.11, 'gram', 'dry'),
('Flour - Rice', 0.0, 0.07, 'gram', 'dry'),
('Flour - Semolina', 10000.0, 0.4, 'gram', 'dry'),
('Flour - Wheat', 0.0, 0.119, 'gram', 'dry'),
('Food Color - Egg Yellow', 330.0, 0.055, 'ml', 'wet'),
('Food Color - Egg Yellow, Powdered', 0.0, 0.21, 'gram', 'dry'),
('Food Color - Red', 950.0, 0.055, 'ml', 'wet'),
('Food Color - Yellow', 0.0, 0.055, 'ml', 'wet'),
('Fresh Coconut', 0.0, 50.0, 'piece', 'fresh'),
('Fries - Shoestring', 26000.0, 0.225, 'gram', 'rtc'),
('Fries - Twister / Loop', 31500.0, 0.312, 'gram', 'rtc'),
('Frozen Tart Shell - Mini', 0.0, 7.0, 'piece', 'dry'),
('Frozen Tart Shell - Regular', 0.0, 14.0, 'piece', 'dry'),
('Fructose', 0.0, 0.17, 'ml', 'wet'),
('Gabi - Fresh', 0.0, 0.31, 'gram', 'fresh'),
('Garam Masala - Powdered', 300.0, 0.81, 'gram', 'dry'),
('Garbanzos - Canned', 0.0, 0.12222222222222222, 'gram', 'wet'),
('Garlic - Fresh', 0.0, 0.18, 'gram', 'dry'),
('Garlic - Powder', 615.0, 0.882, 'gram', 'dry'),
('Garlic Longganisa', 27100.0, 0.37, 'gram', 'pork'),
('Gelatin Powder', 200.0, 5.0, 'gram', 'dry'),
('Gin - Tanqueray', 3000.0, 1.2386666666666666, 'ml', 'beverage'),
('Gin - Zafiro', 19000.0, 0.339, 'ml', 'beverage'),
('Ginger - Fresh', 0.0, 0.18, 'gram', 'fresh'),
('Glucose', 1500.0, 0.47333333333333333, 'ml', 'wet'),
('Glutinous Rice (Malagkit)', 0.0, 0.075, 'gram', 'dry'),
('Gran Marnier 700 ml', 700.0, 2.47, 'ml', 'beverage'),
('Grapes - Green, Seedless', 0.0, 0.82, 'gram', 'fresh'),
('Green Peas - Canned', 0.0, 0.106, 'gram', 'wet'),
('Green Peas - Frozen', 0.0, 0.117, 'gram', 'rtc'),
('Grenadine Syrup', 3000.0, 0.459, 'ml', 'wet'),
('Havana Club White Rum 700ml', 4900.0, 0.719, 'ml', 'beverage'),
('Hazelnut Praline Paste', 0.0, 1.5, 'gram', 'wet'),
('Hoegaarden Regular 500 ml', 0.0, 0.256, 'can', 'beverage'),
('Hoegaarden Rose 500 ml', 0.0, 0.194, 'can', 'beverage'),
('Hoisin Sauce', 1117.0, 0.484, 'ml', 'wet'),
('Honey', 0.0, 0.52, 'ml', 'wet'),
('Hotdog, Tender Juicy', 2000.0, 0.235, 'gram', 'rtc'),
('Ice Cream - Vanilla, Selecta', 0.0, 0.305, 'ml', 'wet'),
('Instant Coffee Ground', 340.0, 1.224, 'gram', 'dry'),
('Instant Dry Yeast', 750.0, 0.46, 'gram', 'dry'),
('Inutak - Small', 0.0, 200.0, 'pack', 'rtc'),
('Italian Seasoning - Dried', 0.0, 2.46, 'gram', 'd'),
('Jose Cuervo Tequila 1 Liter', 0.0, 1.079, 'ml', 'beverage'),
('Kahlua Coffee Liqueur 700 ml', 2800.0, 0.915, 'ml', 'beverage'),
('Kangkong', 0.0, 0.12, 'gram', 'fresh'),
('Kaong - Plain', 12000.0, 0.23, 'jar', 'wet'),
('Kikiam', 0.0, 0.28, 'pack', 'rtc'),
('Knorr Chicken Powder', 6000.0, 0.547, 'gram', 'dry'),
('Knorr Crab and Corn 55g', 1870.0, 1.1272727272727272, 'gram', 'dry'),
('Knorr Demi Glace Powder', 1000.0, 1.378, 'gram', 'dry'),
('Lady Finger', 800.0, 0.95, 'gram', 'dry'),
('Large White Marshmallows', 0.0, 0.4, 'gram', 'dry'),
('Lea & Perrins Worcestershire Sauce', 1773.0, 0.733, 'ml', 'wet'),
('Leche Flan', 0.0, 185.0, 'pack', 'rtc'),
('Lemon - Dried, Sliced', 0.0, 1.0, 'gram', 'dry'),
('Lemon - Fresh', 0.0, 45.0, 'piece', 'fresh'),
('Lemon Grass Powder', 1000.0, 0.388, 'gram', 'dry'),
('Lemongrass - Fresh', 0.0, 0.125, 'gram', 'dry'),
('Lettuce - Green', 0.0, 0.21, 'gram', 'fresh'),
('Liquid Seasoning', 7580.0, 0.09, 'ml', 'wet'),
('Liquid Smoke', 0.0, 3.088, 'ml', 'wet'),
('Loaf Bread per pack', 0.0, 100.0, 'pack', 'gram'),
('Lumpia - Shanghai (12 pcs /pack, per piece)', 504.0, 14.0, 'piece', 'rtc'),
('Lumpia - Siomai (12 pcs /pack, per piece)', 792.0, 14.0, 'piece', 'rtc'),
('Lychee - Canned', 690.0, 0.264, 'gram', 'wet'),
('Macapuno per gram', 18000.0, 0.23, 'jar', 'wet'),
('Magic Sarap', 1600.0, 0.547, 'gram', 'dry'),
('Malunggay Leaves - Fresh', 0.0, 0.21, 'gram', 'dry'),
('Mama Sitas Barbeque Marinade Mix', 1360.0, 0.185, 'gram', 'wet'),
('Mama Sitas Stir Fry Mix', 0.0, 1.45, 'gram', 'dry'),
('Mang Tomas Liver Sauce', 8000.0, 0.119, 'ml', 'wet'),
('Mango - Fresh', 0.0, 0.16, 'gram', 'fresh'),
('Mango - Frozen', 55000.0, 0.33, 'gram', 'frozen'),
('Mango Salsa 750gms', 45750.0, 0.519, 'gram', 'wet'),
('Maple Cured Ham', 3280.0, 0.69, 'gram', 'meat'),
('Maple Syrup', 0.0, 0.3, 'ml', 'wet'),
('Maraschino Cherry', 0.0, 0.678, 'gram', 'dry'),
('Martini Extra Dry Vermouth 1 Liter', 2000.0, 0.818, 'ml', 'beverage'),
('Martini Rosso Vermouth 1 Liter', 2000.0, 0.99, 'ml', 'beverage'),
('Matcha Powder', 2000.0, 0.365, 'gram', 'dry'),
('Mayonnaise - Japanese', 0.0, 0.645, 'ml', 'wet'),
('Mayonnaise - Regular', 5500.0, 0.315, 'ml', 'wet'),
('Milk - Condensed 300ml per ml', 24882.0, 0.25, 'ml', 'wet'),
('Milk - Condensed Ube Flavor 300ml per ml', 0.0, 0.17, 'ml', 'wet'),
('Milk - Evaporated 370ml per ml', 19800.0, 0.176, 'ml', 'wet'),
('Milk - Full Cream / Fresh Milk', 72000.0, 0.095, 'ml', 'wet'),
('Milk - Powdered', 5700.0, 0.308, 'gram', 'dry'),
('Mineral Water 500ml', 365.0, 15.0, 'bottle', 'beverage'),
('Mint - Fresh', 0.0, 0.55, 'gram', 'fresh'),
('Miso Paste', 3950.0, 0.24, 'gram', 'wet'),
('Misua - Dried', 0.0, 0.225, 'gram', 'dry'),
('Mixed Vegetables - Frozen', 10000.0, 0.275, 'gram', 'rtc'),
('Monggo / Mung Beans', 0.0, 0.26, 'gram', 'fresh'),
('Munggo - Bottled (per gram)', 17000.0, 0.23, 'gram', 'wet'),
('Muscovado Sugar', 3000.0, 0.187, 'gram', 'dry'),
('Mushroom - Canned, Sliced', 4000.0, 0.268, 'gram', 'wet'),
('Mushroom - Canned, Whole', 800.0, 0.186, 'gram', 'wet'),
('Mushroom - Fresh, White Button', 0.0, 1.26, 'gram', 'fresh'),
('Mushroom - Shitake, Dried', 0.0, 1.01, 'gram', 'dry'),
('Mushroom - Shitake, Fresh', 0.0, 0.765, 'gram', 'fresh'),
('Mushroom - Tengga Daga', 0.0, 1.01, 'gram', 'dry'),
('Mustard Leaves - Mustasa', 0.0, 0.239, 'gram', 'fresh'),
('Nacho Chips', 0.0, 0.305, 'gram', 'dry'),
('Nata de Coco - Bottled (per gram)', 13000.0, 0.185, 'gram', 'wet'),
('Nestea - Restaurant Blend', 0.0, 1.05, 'gram', 'dry'),
('Nestea Apple Powdered Juice 25 g', 0.0, 1.16, 'gram', 'dry'),
('Nestea Cranberry Powdered Juice 25 g', 0.0, 1.16, 'gram', 'dry'),
('Nestea Lemon Juice - Powdered', 750.0, 1.0, 'gram', 'dry'),
('Nestea Lemonade Cucumber - Powdered', 1680.0, 0.645, 'gram', 'dry'),
('New York Cheesecake - Whole', 0.0, 1350.0, 'piece', 'dessert'),
('Nilupak (per pack)', 0.0, 175.0, 'pack', 'rtc'),
('Niyog - Whole', 0.0, 55.0, 'piece', 'fresh'),
('Noodles - Bihon/Palabok', 0.0, 0.11, 'gram', 'dry'),
('Noodles - Canton', 0.0, 0.17, 'gram', 'dry'),
('Noodles - Miki', 0.0, 0.12, 'gram', 'wet'),
('Noodles - Vermicelli / Sotanghon', 6000.0, 0.189, 'gram', 'dry'),
('Nori Seaweed (small)', 0.0, 5.37, 'gram', 'dry'),
('Nutmeg - Ground', 0.0, 4.351, 'gram', 'dry'),
('Oil - Extra Virgin Olive', 0.0, 0.8, 'ml', 'wet'),
('Oil - Olive', 5000.0, 0.74, 'ml', 'wet'),
('Oil - Sesame', 820.0, 0.55, 'ml', 'wet'),
('Oil - Sunflower', 0.0, 2.2, 'ml', 'wet'),
('Oil - Truffle', 3750.0, 3.42, 'ml', 'wet'),
('Oil - Vegetable (Cooking)', 234000.0, 0.106, 'ml', 'wet'),
('Okra', 0.0, 0.15, 'gram', 'fresh'),
('Olives - Green, Pitted', 2100.0, 0.776, 'ml', 'wet'),
('Onion - Pearl, Pickled', 0.0, 1.05, 'gram', 'wet'),
('Onion - Red', 0.0, 0.2, 'gram', 'fresh'),
('Onion - White', 0.0, 0.18, 'gram', 'fresh'),
('Onion Leeks', 0.0, 0.18, 'gram', 'fresh'),
('Onion Powder', 540.0, 0.86, 'gram', 'dry'),
('Orange - Dried, Sliced', 0.0, 1.2, 'gram', 'dry'),
('Orange - Fresh', 0.0, 55.0, 'piece', 'fresh'),
('Oregano - Fresh', 0.0, 0.35, 'gram', 'fresh'),
('Oregano Leaves - Dried', 1000.0, 1.438, 'gram', 'dry'),
('Oregano Leaves - Ground', 0.0, 2.55, 'gram', 'dry'),
('Ox Tripe', 15080.0, 0.2, 'gram', 'beef'),
('Oyster Sauce', 4054.0, 0.095, 'ml', 'wet'),
('Pampano - Frozen', 30000.0, 0.285, 'gram', 'seafood'),
('Papaya -Fresh', 0.0, 0.2, 'gram', 'fresh'),
('Paprika - Smoked', 0.0, 1.0, 'gram', 'dry'),
('Paprika - Spanish', 1200.0, 2.568, 'gram', 'dry'),
('Parsley - Dried', 0.0, 4.15, 'gram', 'dry'),
('Parsley - Fresh, Curly', 0.0, 0.6, 'gram', 'fresh'),
('Pasta - Fettucine', 0.0, 0.178, 'gram', 'dry'),
('Pasta - Fusili', 3000.0, 0.178, 'gram', 'dry'),
('Pasta - Lasagna', 5000.0, 0.318, 'gram', 'dry'),
('Pasta - Linguine', 12500.0, 0.3, 'gram', 'dry'),
('Pasta - Macaroni', 0.0, 0.141, 'gram', 'dry'),
('Pasta - Penne Rigate', 0.0, 0.189, 'gram', 'dry'),
('Pasta - Spaghetti', 0.0, 0.085, 'gram', 'dry'),
('Peanut - Raw', 0.0, 0.2, 'gram', 'dry'),
('Peanut Butter', 35032.0, 0.345, 'gram', 'wet'),
('Petchay - Baguio', 0.0, 0.08, 'gram', 'fresh'),
('Petchay - Tagalog', 0.0, 0.08, 'gram', 'fresh'),
('Pichi Pichi - Big (per box)', 0.0, 385.0, 'box', 'rtc'),
('Pickle - Relish', 0.0, 0.46, 'ml', 'wet'),
('Pickle - Whole', 0.0, 0.441, 'gram', 'wet'),
('Pineapple - Chunks', 822.0, 0.204, 'gram', 'wet'),
('Pineapple - Fresh', 0.0, 0.18, 'gram', 'fresh'),
('Pineapple - Juice, per ml', 5000.0, 0.132, 'ml', 'wet'),
('Pineapple - Slices', 822.0, 0.312, 'gram', 'wet'),
('Pineapple Juice - 240ml (per can)', 0.0, 40.0, 'can', 'beverage'),
('Pinipig - Raw', 0.0, 0.11, 'gram', 'dry'),
('Pistachio Paste', 0.0, 1.55, 'ml', 'dry'),
('Pixie Luster Duster', 110.0, 12.3, 'gram', 'dry'),
('Popcorn - Raw', 0.0, 0.075, 'gram', 'dry'),
('Pork - Belly - BISO', 149701.0, 0.43, 'gram', 'pork'),
('Pork - Belly - BLSL', 5530.0, 0.43, 'gram', 'pork'),
('Pork - Belly - Sliced', 0.0, 0.43, 'gram', 'pork'),
('Pork - Bulaklak', 0.0, 0.15, 'gram', 'pork'),
('Pork - Chops', 0.0, 0.32, 'gram', 'pork'),
('Pork - Fats', 0.0, 0.22, 'gram', 'pork'),
('Pork - Ground', 0.0, 0.275, 'gram', 'pork'),
('Pork - Jowls', 0.0, 0.28, 'gram', 'pork'),
('Pork - Kasim / Shoulder', 8370.0, 0.28, 'gram', 'pork'),
('Pork - Liver', 0.0, 0.3, 'gram', 'pork'),
('Pork - Pata Front', 0.0, 0.225, 'gram', 'pork'),
('Pork - Smoked Bacon', 6000.0, 0.767, 'gram', 'pork'),
('Pork - Thick - Bacon Picnic', 5215.0, 1.014, 'gram', 'pork'),
('Potato - Fresh', 0.0, 0.14, 'gram', 'fresh'),
('Potato Starch', 0.0, 0.443, 'gram', 'dry'),
("Powdered / Confectioner's Sugar", 6000.0, 0.296, 'gram', 'dry'),
('Proseco 750 ml', 3000.0, 1.079, 'ml', 'beverage'),
('Puff Pastry Sheet (pack of 5 sheets)', 0.0, 1125.0, 'pack', 'dry'),
('Puratos - Fudgy Brownie Mix (per gram)', 0.0, 0.53, 'gram', 'dry'),
('Puratos - Red Velvet Cake Mix 1kg (per gram)', 11000.0, 0.53, 'gram', 'dry'),
('Puratos - Tegral Vanilla Sponge (per gram)', 0.0, 0.53, 'gram', 'dry'),
('Puratos - Ube Cake Mix (per gram)', 2000.0, 0.53, 'gram', 'dry'),
('Puratos- Chocolate Cake Mix (per gram)', 7000.0, 0.53, 'gram', 'dry'),
('Quail Egg', 0.0, 4.0, 'piece', 'chicken'),
('Raddish - Fresh', 0.0, 0.08, 'gram', 'fresh'),
('Raisins', 0.0, 0.565, 'gram', 'dry'),
('Red Ice Tea - Powdered', 0.0, 0.84, 'gram', 'dry'),
('Reno Liver Spread', 4830.0, 0.296, 'ml', 'wet'),
('Rice - Arborio', 1000.0, 1.02, 'gram', 'dry'),
('Rice - White', 10000.0, 0.068, 'gram', 'dry'),
('Rice Wine', 0.0, 0.24, 'ml', 'wet'),
('Rite n Lite - Cucumber', 0.0, 35.0, 'piece', 'beverage'),
('Rootbeer (per ml)', 0.0, 0.5, 'ml', 'wet'),
('Rosemary Leaves - Dried', 0.0, 1.779, 'gram', 'dry'),
('Sago - Bottled (per gram)', 9000.0, 0.13, 'gram', 'wet'),
('Salmon - Frozen Slab', 15100.0, 1.0, 'gram', 'seafood'),
('Salt - Iodized', 3000.0, 0.04, 'gram', 'dry'),
('Salt - Rock', 0.0, 0.055, 'gram', 'dry'),
('San Miguel Light 330 ml', 48.0, 60.0, 'can', 'beverage'),
('San Miguel Pale Pilsen 330 ml', 96.0, 60.0, 'can', 'beverage'),
('Sausage - Cocktail Hungarian', 11440.0, 0.59, 'gram', 'pork'),
('Sayote - Fresh', 0.0, 0.08, 'gram', 'fresh'),
('Sesame Seeds', 0.0, 0.025, 'gram', 'dry'),
('Sesame Seeds - Black', 120.0, 0.39, 'gram', 'dry'),
('Shotts - Passion Fruit', 0.0, 0.91, 'ml', 'beverage'),
('Shotts - Rose', 0.0, 0.91, 'ml', 'beverage'),
('Shotts - Black Tea & Peach', 5000.0, 0.91, 'ml', 'beverage'),
('Shotts - Butterscotch', 4000.0, 0.71, 'ml', 'beverage'),
('Shotts - Cane Sugar', 0.0, 0.71, 'ml', 'beverage'),
('Shotts - Caramel', 4000.0, 0.71, 'ml', 'beverage'),
('Shotts - Dark Chocolate', 0.0, 0.71, 'ml', 'beverage'),
('Shotts - Flamed Orange', 3000.0, 0.91, 'ml', 'beverage'),
('Shotts - Hazelnut', 4000.0, 0.71, 'ml', 'beverage'),
('Shotts - Irish Cream', 0.0, 0.71, 'ml', 'beverage'),
('Shotts - Lemon Ginger & Honey', 7000.0, 0.91, 'ml', 'beverage'),
('Shotts - Lychee', 4000.0, 0.91, 'ml', 'beverage'),
('Shotts - Macadamia', 4000.0, 0.71, 'ml', 'beverage'),
('Shotts - Salted Caramel', 4000.0, 0.71, 'ml', 'beverage'),
('Shotts - Sticky Strawberry', 13000.0, 0.91, 'ml', 'beverage'),
('Shotts - Strawberry', 0.0, 0.91, 'ml', 'beverage'),
('Shotts - Tahitian Lime', 5000.0, 0.91, 'ml', 'beverage'),
('Shotts - Triple Peach', 3000.0, 0.91, 'ml', 'beverage'),
('Shotts - Vanilla', 5000.0, 0.91, 'ml', 'beverage'),
('Shotts - White Chocolate', 4000.0, 0.71, 'ml', 'beverage'),
('Shrimp - Fresh, Whole', 0.0, 0.55, 'gram', 'seafood'),
('Shrimp - Peeled, Frozen', 19000.0, 0.48, 'gram', 'seafood'),
('Shrimp Paste', 370.0, 0.767, 'gram', 'wet'),
('Sinamak 750 ml', 0.0, 0.267, 'ml', 'wet'),
('Singkamas', 0.0, 0.07, 'gram', 'fresh'),
('Sinigang Mix - Gabi', 3300.0, 0.796, 'gram', 'dry'),
('Sinigang Mix - Miso', 1564.0, 1.087, 'gram', 'dry'),
('Sinigang Mix - Original', 0.0, 1.091, 'gram', 'dry'),
('Sitaw', 0.0, 0.13, 'gram', 'fresh'),
('Sitsaro', 0.0, 0.28, 'gram', 'fresh'),
('Sliced Cheese per gram', 0.0, 1.0, 'gram', 'rtc'),
('Sliced Jalapenos', 335.0, 0.465, 'gram', 'wet'),
('Smirnoff Vodka 700ml', 700.0, 1.212, 'ml', 'beverage'),
('Soda Water 320 ml (per ml)', 3200.0, 0.125, 'ml', 'beverage'),
('Soy Sauce - Dark', 4500.0, 0.186, 'ml', 'wet'),
('Soy Sauce - Kikkoman', 5000.0, 0.292, 'ml', 'wet'),
('Soy Sauce - Regular', 15600.0, 0.105, 'ml', 'wet'),
('Spaghetti Sauce - Filipino Style', 0.0, 0.13, 'ml', 'wet'),
('Spam Regular 340 g (per gram)', 2040.0, 0.671, 'gram', 'rtc'),
('Spring Onion', 0.0, 0.61, 'gram', 'fresh'),
('Sprite 320 ml', 72.0, 37.0, 'can', 'beverage'),
('Squash', 0.0, 0.08, 'gram', 'fresh'),
('Sriracha', 0.0, 0.655, 'ml', 'wet'),
('Star Anise', 100.0, 1.79, 'gram', 'dry'),
('Star Margarine', 1000.0, 0.36, 'gram', 'wet'),
('Strawberry - Fruit Jam', 0.0, 0.5, 'ml', 'wet'),
('Strawberry Syrup - Dionysus', 0.0, 1.07, 'ml', 'wet'),
('Sugar - Brown', 30000.0, 0.11, 'gram', 'dry'),
('Sugar - Pearl', 0.0, 0.505, 'gram', 'wet'),
('Sugar - White', 25000.0, 0.138, 'gram', 'dry'),
('Sweet & Chili Dipping Sauce', 5520.0, 0.261, 'ml', 'wet'),
('Sweet & Sour Sauce', 4140.0, 0.384, 'ml', 'wet'),
('Sweet Basil - Fresh', 0.0, 0.65, 'gram', 'fresh'),
('Sweet Corn - Fresh', 0.0, 0.07, 'gram', 'fresh'),
('Sweet Potato', 0.0, 0.09, 'gram', 'fresh'),
('Tabasco Sauce', 240.0, 1.886, 'ml', 'wet'),
('Tamarind  Paste', 908.0, 0.33, 'gram', 'wet'),
('Tanduay Rum - 1 Liter (per ml)', 0.0, 0.24, 'ml', 'beverage'),
('Tang Orange Powdered Juice 19g (per gram)', 375.0, 0.95, 'gram', 'dry'),
('Tanglad / Lemongrass - Fresh', 0.0, 0.25, 'gram', 'fresh'),
('Tapioca Pearl, Small, Raw', 0.0, 0.15, 'gram', 'dry'),
('Taro/Gabi', 0.0, 0.31, 'gram', 'fresh'),
('Tea - Chamomile', 0.0, 1.15, 'gram', 'dry'),
('Tea - Green Jasmine', 200.0, 1.05, 'gram', 'dry'),
('Tea - Hibiscus', 200.0, 2.05, 'gram', 'dry'),
('Teriyaki Sauce', 0.0, 0.27, 'ml', 'wet'),
('Thyme Leaves - Dried', 500.0, 1.572, 'gram', 'dry'),
('Tokwa - Fresh', 0.0, 0.135, 'gram', 'fresh'),
('Tomato Ketchup', 29931.0, 0.199, 'ml', 'wet'),
('Tomato Paste', 0.0, 0.3, 'gram', 'wet'),
('Tomato Sauce - Filipino Style', 0.0, 0.13, 'ml', 'wet'),
('Tomatoes - Canned, Diced', 0.0, 0.25, 'gram', 'wet'),
('Tomatoes - Canned, Whole peeled', 0.0, 0.177, 'gram', 'wet'),
('Tomatoes - Fresh', 0.0, 0.12, 'gram', 'fresh'),
('Tomatoes - Sundried', 0.0, 1.347, 'gram', 'wet'),
('Tonic Water 320 ml', 0.0, 45.0, 'can', 'beverage'),
('Torani White Chocolate Sauce 1.89 Liter (per ml)', 0.0, 0.847, 'ml', 'wet'),
('Triple Sec 1 Liter', 0.0, 0.5, 'ml', 'beverage'),
('Tuna - Belly, Frozen', 77610.0, 0.5, 'gram', 'seafood'),
('Tuna - Yellow Fin, Fresh', 0.0, 0.35, 'gram', 'seafood'),
('Tuna Flakes In Oil 170g per gram', 170.0, 0.235, 'gram', 'wet'),
('Turmeric - Ground', 0.0, 0.65, 'gram', 'dry'),
('Ube - Don Benitos (per pack)', 0.0, 220.0, 'pack', 'rtc'),
('Ube Extract', 2400.0, 1.723, 'ml', 'wet'),
('Ube Powder', 0.0, 0.45, 'gram', 'dry'),
('Ube Syrup - Dionysus', 0.0, 2.312, 'ml', 'wet'),
('Vanilla Extract', 1425.0, 1.053, 'ml', 'wet'),
('Vanilla Powder', 0.0, 0.355, 'gram', 'dry'),
('Vetsin', 1000.0, 0.25, 'gram', 'dry'),
('Vinegar - Apple Cider', 0.0, 0.529, 'ml', 'wet'),
('Vinegar - Balsamic', 0.0, 0.5, 'ml', 'wet'),
('Vinegar - Cane', 3785.0, 0.053, 'ml', 'wet'),
('Vinegar - Red Wine', 0.0, 1.2, 'ml', 'wet'),
('Vinegar - White', 3785.0, 0.05, 'ml', 'wet'),
('Walnuts - Whole, Raw', 1000.0, 0.8, 'gram', 'dry'),
('Wansoy/Cilantro - Fresh', 0.0, 0.46, 'gram', 'fresh'),
('Watermelon - Fresh', 0.0, 0.06, 'gram', 'fresh'),
('Whipped Topping - Ambiante', 2000.0, 0.32, 'pack', 'wet'),
('Whipped Topping - Everwhip', 0.0, 0.25, 'ml', 'wet'),
('Whiskey - Andy Player', 0.0, 0.52, 'ml', 'beverage'),
('Whiskey - Irishman', 700.0, 3.358, 'ml', 'beverage'),
('Whiskey - Jack Daniel', 7000.0, 2.1, 'ml', 'beverage'),
('Whiskey - Jim Beam', 3000.0, 1.35, 'ml', 'beverage'),
('Whiskey - Johnny Walker', 0.0, 1.65, 'ml', 'beverage'),
('White Chocolate - Bar', 0.0, 0.39, 'gram', 'dry'),
('White Chocolate - Droplets', 0.0, 0.37, 'gram', 'dry'),
('White Pepper - Ground', 0.0, 0.39, 'gram', 'dry'),
('White Rum 700 ml', 0.0, 1.439, 'ml', 'beverage'),
('White Sugar (Sachet) - 100pcs (per piece)', 500.0, 2.0, 'piece', 'beverage'),
('Wine - Red - 750 ml', 22.0, 555.0, 'bottle', 'beverage'),
('Wine - Rose - 750ml', 18.0, 555.0, 'bottle', 'beverage'),
('Wine - Sangria', 4.0, 1350.0, 'ml', 'beverage'),
('Wine - White 750 ml', 11.0, 450.0, 'bottle', 'beverage'),
('Young Corn - Canned', 0.0, 0.3, 'gram', 'wet'),
('Mussels - Frozen', 3000.0, 0.25, 'gram', 'frozen'),
('Squid - Frozen', 0.0, 0.38, 'gram', 'frozen'),
('Bukchoy - Fresh', 0.0, 0.22, 'gram', 'fresh'),
('Johnny Walker Black Label (per bottle)', 5.0, 1200.0, 'bottle', 'beverage'),
('Molo Wrapper', 840.0, 0.19, 'gram', 'fresh'),
('Gochujang Paste', 0.0, 0.824, 'gram', 'wet'),
('Cassava Cake', 0.0, 170.0, 'pack', 'gram'),
('Suman (per gram)', 0.0, 0.15, 'gram', 'fresh'),
('Panutsa', 0.0, 0.08, 'gram', 'dry'),
('San Miguel Apple', 0.0, 65.0, 'can', 'beverage'),
('Coconut Rum - Orchid', 0.0, 0.3, 'ml', 'beverage'),
('Lime Juice - Orchid', 0.0, 0.16, 'ml', 'beverage'),
('Tequila - El Hombre', 0.0, 0.572, 'ml', 'beverage'),
('GSM Blue', 0.0, 0.286, 'ml', 'beverage'),
('Danggit', 3000.0, 1.7, 'gram', 'dry'),
('Curry Powder', 0.0, 0.32, 'gram', 'dry'),
('Red Horse', 0.0, 83.0, 'can', 'beverage'),
('Sweet Ham', 0.0, 0.6, 'gram', 'frozen'),
('Cheedar Cheese Powder', 0.0, 0.275, 'gram', 'dry'),
('Barbeque Powder', 0.0, 0.275, 'gram', 'dry'),
('Kropek', 0.0, 0.9, 'gram', 'dry'),
('Mixed Cheese', 0.0, 0.6, 'gram', 'dry'),
('Rosemary Leaves - Fresh', 0.0, 0.75, 'gram', 'fresh'),
('Graham Crackers', 0.0, 0.3, 'gram', 'dry'),
('Milk - Coconut', 0.0, 0.185, 'gram', 'dry'),
('Curing / Pink Salt / Prague Powder', 0.0, 0.12, 'gram', 'dry'),
('Cloves - Whole', 0.0, 1.265, 'gram', 'dry'),
('Curry - Green', 0.0, 1.84, 'gram', 'rtc'),
('Cascina Italian Pesto per gram', 0.0, 0.7481481481481481, 'gram', 'rtc'),
('Cassava Starch', 0.0, 0.15, 'gram', 'dry'),
('Corn Syrup', 0.0, 0.472, 'ml', 'rtc'),
('Food Color - Pink', 0.0, 1.6, 'ml', 'wet'),
('Food Color - Green', 0.0, 1.6, 'ml', 'wet'),
('Food Color - Brown', 0.0, 1.6, 'ml', 'wet'),
('Isomalt', 0.0, 2.864, 'gram', 'dry'),
('Lemon Extract', 0.0, 1.579, 'ml', 'wet'),
('Lime - Dried, Sliced', 0.0, 1.94, 'gram', 'dry'),
('Fruit Cocktail', 0.0, 0.117, 'grams', 'dry'),
('Gulaman - Buko Pandan', 275.0, 1.2, 'grams', 'dry'),
('Sour Cream Powder', 0.0, 0.275, 'gram', 'dry'),
('Coconut Meat - per piece', 0.0, 45.0, 'pcs', 'fresh'),
('Stevia sweetener - per piece', 500.0, 3.0, 'piece', 'dry'),
('Sago - Raw', 0.0, 0.22, 'gram', 'dry'),
('Milk - Soy (Unsweetened)', 0.0, 0.17, 'gram', 'wet'),
('Dragon Fruit', 0.0, 0.499, 'grams', 'fresh'),
('Cashew Nuts', 0.0, 0.4, 'gram', 'fresh'),
('Nata bits- 1000g', 5000.0, 0.36, 'gram', 'wet'),
('Togue', 0.0, 0.2, 'gram', 'fresh'),
('Pancake Mix- Maya', 0.0, 0.25, 'gram', 'dry'),
('Gulaman - Mango', 600.0, 0.64, 'gram', 'dry'),
('Shotts - Coconut', 0.0, 0.91, 'ml', 'beverage'),
('Langka', 0.0, 0.085, 'gram', 'dry'),
('Sampaloc', 0.0, 0.15, 'gram', 'dry'),
('Masa Picha -Ham & Cheese', 0.0, 95.0, 'piece', 'dry'),
('Masa Picha -Hawaiian', 0.0, 100.0, 'piece', 'dry'),
('Masa Picha -All Meat', 0.0, 135.0, 'piece', 'dry'),
('Ice Candy- Mango', 0.0, 20.0, 'piece', 'frozen'),
('Shotts- Pink Grape Fruit', 0.0, 0.9, 'ml', 'wet'),
('Milk- Oatside', 0.0, 0.19, 'ml', 'wet'),
('Biscoff - Sauce', 0.0, 1.05, 'ml', 'wet'),
('Caramel -Sauce', 0.0, 0.33, 'ml', 'wet'),
('Cassava Cups - Cheese', 0.0, 40.0, 'pcs', 'rtc'),
('Cassava Cups - Corn', 0.0, 40.0, 'pcs', 'rtc'),
('Cassava Cups - Macapuno', 0.0, 40.0, 'pcs', 'rtc'),
('Salted Egg Powder', 0.0, 1.65, 'g', 'dry'),
('TOTAL WORTH OF STOCKS', 0.0, 0.0, '', ''),
]

# ── Google Sheets ──────────────────────────────────────────────────────────────
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet():
    return get_client().open_by_key(SPREADSHEET_ID)

def ensure_sheet(ss, name, headers):
    try:
        ws = ss.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=5000, cols=len(headers)+2)
        ws.append_row(headers)
    return ws

@st.cache_data(ttl=30)
def load_items(_ss_id):
    ss = get_spreadsheet()
    ws = ensure_sheet(ss, ITEMS_SHEET, ITEMS_HEADERS)
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=ITEMS_HEADERS)

@st.cache_data(ttl=30)
def load_log(_ss_id):
    ss = get_spreadsheet()
    ws = ensure_sheet(ss, LOG_SHEET, LOG_HEADERS)
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=LOG_HEADERS)

def invalidate_cache():
    load_items.clear()
    load_log.clear()

def num(val):
    try: return float(val or 0)
    except: return 0.0

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='background:linear-gradient(175deg,#1A2E1A 0%,#0E1C0E 60%,#080F08 100%);
                padding:24px 20px 18px;border-bottom:1px solid #2A3828;text-align:center;'>
        <div style='font-family:Cormorant Garamond,serif;font-size:1.5rem;font-weight:700;color:#8CAF7A;letter-spacing:2px;'>SERVANDO</div>
        <div style='font-size:0.56rem;letter-spacing:3px;text-transform:uppercase;color:#4A6B3E;margin-top:5px;'>Main Warehouse Inventory</div>
    </div>
    <div style='font-size:0.56rem;font-weight:600;letter-spacing:3px;text-transform:uppercase;
                color:#4A6B3E;padding:14px 12px 5px;border-bottom:1px solid #1E2E1C;margin-bottom:6px;'>
        Navigation
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("", [
        "📊 Dashboard",
        "📥 Incoming Deliveries",
        "📋 Purchase Orders",
        "🔧 Stock Adjustment",
        "📆 Generate Report",
        "🔍 Item History",
        "📦 Items Master",
        "⬇️ Export to Excel",
        "⚙️ Setup",
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown(f"<div style='font-size:0.62rem;letter-spacing:1.5px;text-transform:uppercase;color:#3A5238;'>{date.today().strftime('%B %d, %Y')}</div>", unsafe_allow_html=True)

# ── Connect ────────────────────────────────────────────────────────────────────
try:
    ss = get_spreadsheet()
except Exception as e:
    st.error(f"❌ Cannot connect to Google Sheets: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.markdown('<div class="main-header"><h1>📊 Dashboard</h1><p>Live stock overview</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    log_df   = load_log(SPREADSHEET_ID)

    if items_df.empty:
        st.info("No items yet. Go to ⚙️ Setup to initialize.")
        st.stop()

    active = items_df[items_df["ACTIVE"] == "YES"].copy()

    # Compute current stock per item from log
    def compute_stock(item_name, beginning):
        if log_df.empty:
            return beginning
        ilog = log_df[log_df["ITEM"] == item_name]
        if ilog.empty:
            return beginning
        add_in   = ilog["ADD_IN"].apply(num).sum()
        over     = ilog["OVER"].apply(num).sum()
        rest     = ilog["RESTAURANT"].apply(num).sum()
        banq     = ilog["BANQUET"].apply(num).sum()
        cafe     = ilog["CAFE"].apply(num).sum()
        bar      = ilog["BAR"].apply(num).sum()
        others   = ilog["OTHERS"].apply(num).sum()
        spoil    = ilog["SPOILAGE"].apply(num).sum()
        return beginning + add_in + over - rest - banq - cafe - bar - others - spoil

    active["CURRENT_STOCK"] = active.apply(lambda r: compute_stock(r["ITEM"], num(r["BEGINNING_STOCKS"])), axis=1)
    active["TOTAL_WORTH"]   = active["CURRENT_STOCK"] * active["UNIT COST"].apply(num)

    total_worth  = active["TOTAL_WORTH"].sum()
    total_items  = len(active)
    low_stock    = len(active[active["CURRENT_STOCK"] <= 0])
    total_txns   = len(log_df) if not log_df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Total Items", total_items)
    with c2: st.metric("Total Stock Worth", f"₱{total_worth:,.2f}")
    with c3: st.metric("Zero / Negative Stock", low_stock)
    with c4: st.metric("Total Transactions", total_txns)

    st.markdown("---")

    # Recent transactions
    if not log_df.empty:
        st.markdown('<div class="section-title">Recent Transactions</div>', unsafe_allow_html=True)
        recent = log_df.sort_values("TIMESTAMP", ascending=False).head(10)
        for _, row in recent.iterrows():
            txn = row.get("TXN_TYPE","")
            ref = row.get("REF_NUMBER","")
            icon = "📥" if txn == "DELIVERY" else ("📋" if txn == "PO" else "🔧")
            ref_span = f'<span style="color:#5A7A52;">({ref})</span>' if ref else ""
            st.markdown(f'<div class="log-entry">{icon} <strong>{row.get("ITEM","")}</strong> &nbsp;·&nbsp; {txn} {ref_span} &nbsp;·&nbsp; 👤 {row.get("STAFF","")}&nbsp;·&nbsp; 📅 {row.get("DATE","")}</div>', unsafe_allow_html=True)

    # Low / zero stock warning
    zero = active[active["CURRENT_STOCK"] <= 0]
    if not zero.empty:
        st.markdown('<div class="section-title" style="color:#CC6A6A;">⚠️ Zero or Negative Stock Items</div>', unsafe_allow_html=True)
        st.dataframe(zero[["ITEM","UNIT OF MEASURE","CATEGORY","CURRENT_STOCK"]].sort_values("ITEM"),
                     use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DELIVERIES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📥 Incoming Deliveries":
    st.markdown('<div class="main-header"><h1>📥 Incoming Deliveries</h1><p>Log incoming stock from suppliers</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    active_items = items_df[items_df["ACTIVE"]=="YES"]["ITEM"].tolist() if not items_df.empty else []

    if not active_items:
        st.warning("No items found. Go to Setup first.")
        st.stop()

    if "delivery_cart" not in st.session_state:
        st.session_state.delivery_cart = []

    # Header
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Delivery Details</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: delivery_date = st.date_input("📅 Delivery Date", value=date.today(), key="del_date")
    with c2: staff_name = st.text_input("👤 Received By", placeholder="Your name", key="del_staff")
    st.markdown('</div>', unsafe_allow_html=True)

    # Add item
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Add Item to Delivery</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([3,1])
    with c1: sel_item = st.selectbox("Select Item", active_items, key="del_item")
    with c2: qty = st.number_input("Quantity", min_value=0.01, step=0.01, format="%.2f", key="del_qty")

    if sel_item:
        info = items_df[items_df["ITEM"]==sel_item].iloc[0]
        st.markdown(f'<div class="info-box">📦 <strong>{sel_item}</strong> &nbsp;|&nbsp; {info["UNIT OF MEASURE"]} &nbsp;|&nbsp; ₱{num(info["UNIT COST"]):.4f}/unit &nbsp;|&nbsp; {info["CATEGORY"].upper()}</div>', unsafe_allow_html=True)

    if st.button("➕ Add to Delivery", use_container_width=True):
        if sel_item and qty > 0:
            st.session_state.delivery_cart.append({
                "item": sel_item,
                "qty": qty,
                "unit": info["UNIT OF MEASURE"],
                "cost": num(info["UNIT COST"]),
                "notes": ""
            })
            st.success(f"✅ {sel_item} added.")
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Cart
    if st.session_state.delivery_cart:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">Delivery Cart — {len(st.session_state.delivery_cart)} item(s)</div>', unsafe_allow_html=True)
        total_val = 0
        for idx, it in enumerate(st.session_state.delivery_cart):
            val = it["qty"] * it["cost"]
            total_val += val
            col_info, col_del = st.columns([5,1])
            with col_info:
                notes_str = f' &nbsp;·&nbsp; {it["notes"]}' if it['notes'] else ''
                st.markdown(f'<div class="po-item-row"><strong>{idx+1}. {it["item"]}</strong> &nbsp;|&nbsp; <span style="color:#8CAF7A;">{it["qty"]:,.2f} {it["unit"]}</span> &nbsp;|&nbsp; ₱{val:,.2f}{notes_str}</div>', unsafe_allow_html=True)
            with col_del:
                if st.button("🗑️", key=f"del_rm_{idx}"):
                    st.session_state.delivery_cart.pop(idx); st.rerun()

        st.markdown(f'<div class="info-box" style="text-align:right;">Total Delivery Value: <strong>₱{total_val:,.2f}</strong></div>', unsafe_allow_html=True)
        st.markdown("---")
        c1, c2 = st.columns([1,2])
        with c1:
            if st.button("🗑️ Clear", use_container_width=True):
                st.session_state.delivery_cart = []; st.rerun()
        with c2:
            if st.button("💾 Submit Delivery", type="primary", use_container_width=True):
                if not staff_name.strip():
                    st.error("Please enter your name.")
                else:
                    ref = f"DEL-{delivery_date.strftime('%Y%m%d')}-{datetime.now().strftime('%H%M%S')}"
                    log_ws = ensure_sheet(ss, LOG_SHEET, LOG_HEADERS)
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    rows = []
                    for it in st.session_state.delivery_cart:
                        rows.append([ts, delivery_date.strftime("%Y-%m-%d"),
                                     delivery_date.strftime("%b %Y").upper(),
                                     it["item"], staff_name.strip(), "DELIVERY", ref,
                                     it["qty"], 0, 0, 0, 0, 0, 0, 0, it["notes"]])
                    log_ws.append_rows(rows)
                    invalidate_cache()
                    n = len(st.session_state.delivery_cart)
                    st.session_state.delivery_cart = []
                    st.success(f"✅ Delivery **{ref}** submitted! {n} item(s) logged.")
                    st.balloons()
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: PURCHASE ORDERS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Purchase Orders":
    st.markdown('<div class="main-header"><h1>📋 Purchase Orders</h1><p>Release stock to departments</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    active_items = items_df[items_df["ACTIVE"]=="YES"]["ITEM"].tolist() if not items_df.empty else []

    if not active_items:
        st.warning("No items found. Go to Setup first.")
        st.stop()

    if "po_cart" not in st.session_state:
        st.session_state.po_cart = []

    # Header
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">PO Details</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: po_date = st.date_input("📅 Date", value=date.today(), key="po_date")
    with c2: po_staff = st.text_input("👤 Prepared By", placeholder="Your name", key="po_staff")
    with c3:
        dept_choice = st.selectbox("🏢 Department", DEPARTMENTS + ["Others (Specify)"], key="po_dept")
    if dept_choice == "Others (Specify)":
        dept_specify = st.text_input("Specify department", key="po_dept_other")
        department = dept_specify.strip() if dept_specify.strip() else "Others"
    else:
        department = dept_choice
    st.markdown('</div>', unsafe_allow_html=True)

    # Add item
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Add Item to PO</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([3,1])
    with c1: po_item = st.selectbox("Select Item", active_items, key="po_item")
    with c2: po_qty = st.number_input("Quantity", min_value=0.01, step=0.01, format="%.2f", key="po_qty")

    if po_item:
        info = items_df[items_df["ITEM"]==po_item].iloc[0]
        st.markdown(f'<div class="info-box">📦 <strong>{po_item}</strong> &nbsp;|&nbsp; {info["UNIT OF MEASURE"]} &nbsp;|&nbsp; ₱{num(info["UNIT COST"]):.4f}/unit &nbsp;|&nbsp; {info["CATEGORY"].upper()}</div>', unsafe_allow_html=True)

    if st.button("➕ Add to PO", use_container_width=True):
        if po_item and po_qty > 0:
            st.session_state.po_cart.append({"item": po_item, "qty": po_qty, "notes": "",
                                              "unit": info["UNIT OF MEASURE"], "cost": num(info["UNIT COST"]),
                                              "dept": department})
            st.success(f"✅ {po_item} added.")
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Cart
    if st.session_state.po_cart:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">PO Review — {len(st.session_state.po_cart)} item(s) → {department}</div>', unsafe_allow_html=True)
        total_val = 0
        for idx, it in enumerate(st.session_state.po_cart):
            val = it["qty"] * it["cost"]
            total_val += val
            c_info, c_del = st.columns([5,1])
            with c_info:
                notes_str = f' &nbsp;·&nbsp; {it["notes"]}' if it['notes'] else ''
                st.markdown(f'<div class="po-item-row"><strong>{idx+1}. {it["item"]}</strong> &nbsp;|&nbsp; <span style="color:#8CAF7A;">{it["qty"]:,.2f} {it["unit"]}</span> &nbsp;|&nbsp; ₱{val:,.2f}{notes_str}</div>', unsafe_allow_html=True)
            with c_del:
                if st.button("🗑️", key=f"po_rm_{idx}"):
                    st.session_state.po_cart.pop(idx); st.rerun()

        st.markdown(f'<div class="info-box" style="text-align:right;">Total PO Value: <strong>₱{total_val:,.2f}</strong></div>', unsafe_allow_html=True)
        st.markdown("---")
        c1, c2 = st.columns([1,2])
        with c1:
            if st.button("🗑️ Clear PO", use_container_width=True):
                st.session_state.po_cart = []; st.rerun()
        with c2:
            if st.button("💾 Submit PO", type="primary", use_container_width=True):
                if not po_staff.strip():
                    st.error("Please enter your name.")
                else:
                    po_ref = f"PO-{po_date.strftime('%Y%m%d')}-{department[:3].upper()}-{datetime.now().strftime('%H%M%S')}"
                    log_ws = ensure_sheet(ss, LOG_SHEET, LOG_HEADERS)
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    dept_map = {"Restaurant":"RESTAURANT","Banquet":"BANQUET","Café":"CAFE","Bar":"BAR"}
                    rows = []
                    for it in st.session_state.po_cart:
                        dept_col = it["dept"]
                        r = [ts, po_date.strftime("%Y-%m-%d"), po_date.strftime("%b %Y").upper(),
                             it["item"], po_staff.strip(), "PO", po_ref,
                             0, 0, 0, 0, 0, 0, 0, 0, it["notes"]]
                        # Set the right dept column
                        dept_indices = {"Restaurant":9,"Banquet":10,"Café":11,"Bar":12}
                        if dept_col in dept_indices:
                            r[dept_indices[dept_col]] = it["qty"]
                        else:
                            r[13] = it["qty"]  # Others
                        rows.append(r)
                    log_ws.append_rows(rows)
                    invalidate_cache()
                    n = len(st.session_state.po_cart)
                    st.session_state.po_cart = []
                    st.success(f"✅ PO **{po_ref}** submitted! {n} item(s) released to **{department}**.")
                    st.balloons()
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: STOCK ADJUSTMENT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔧 Stock Adjustment":
    st.markdown('<div class="main-header"><h1>🔧 Stock Adjustment</h1><p>Correct stock levels · Spoilage · Count adjustments</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    active_items = items_df[items_df["ACTIVE"]=="YES"]["ITEM"].tolist() if not items_df.empty else []

    if not active_items:
        st.warning("No items found."); st.stop()

    if "adj_cart" not in st.session_state:
        st.session_state.adj_cart = []

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Adjustment Details</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: adj_date = st.date_input("📅 Date", value=date.today(), key="adj_date")
    with c2: adj_staff = st.text_input("👤 Adjusted By", placeholder="Your name", key="adj_staff")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Add Adjustment</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([3,1.5,1])
    with c1: adj_item = st.selectbox("Select Item", active_items, key="adj_item")
    with c2:
        adj_type = st.selectbox("Type", ["Over (Add +)", "Spoilage (Remove −)"], key="adj_type")
    with c3: adj_qty = st.number_input("Qty", min_value=0.01, step=0.01, format="%.2f", key="adj_qty")
    adj_notes = st.text_input("Reason (required)", placeholder="e.g. physical count variance, expired goods...", key="adj_notes")

    if adj_item:
        info = items_df[items_df["ITEM"]==adj_item].iloc[0]
        st.markdown(f'<div class="info-box">📦 <strong>{adj_item}</strong> &nbsp;|&nbsp; {info["UNIT OF MEASURE"]} &nbsp;|&nbsp; ₱{num(info["UNIT COST"]):.4f}/unit</div>', unsafe_allow_html=True)

    if st.button("➕ Add Adjustment", use_container_width=True):
        if not adj_notes.strip():
            st.error("Please provide a reason for the adjustment.")
        elif adj_item and adj_qty > 0:
            st.session_state.adj_cart.append({
                "item": adj_item, "type": adj_type, "qty": adj_qty,
                "notes": adj_notes, "unit": info["UNIT OF MEASURE"], "cost": num(info["UNIT COST"])
            })
            st.success(f"✅ {adj_item} added.")
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.adj_cart:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">Adjustments to Submit — {len(st.session_state.adj_cart)} item(s)</div>', unsafe_allow_html=True)
        for idx, it in enumerate(st.session_state.adj_cart):
            sign = "➕" if "Over" in it["type"] else "➖"
            c_i, c_d = st.columns([5,1])
            with c_i:
                st.markdown(f'<div class="po-item-row">{sign} <strong>{it["item"]}</strong> &nbsp;|&nbsp; {it["qty"]:,.2f} {it["unit"]} &nbsp;|&nbsp; <span style="color:#5A7A52;">{it["type"]}</span> &nbsp;|&nbsp; {it["notes"]}</div>', unsafe_allow_html=True)
            with c_d:
                if st.button("🗑️", key=f"adj_rm_{idx}"):
                    st.session_state.adj_cart.pop(idx); st.rerun()

        st.markdown("---")
        c1, c2 = st.columns([1,2])
        with c1:
            if st.button("🗑️ Clear", use_container_width=True):
                st.session_state.adj_cart = []; st.rerun()
        with c2:
            if st.button("💾 Submit Adjustments", type="primary", use_container_width=True):
                if not adj_staff.strip():
                    st.error("Please enter your name.")
                else:
                    ref = f"ADJ-{adj_date.strftime('%Y%m%d')}-{datetime.now().strftime('%H%M%S')}"
                    log_ws = ensure_sheet(ss, LOG_SHEET, LOG_HEADERS)
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    rows = []
                    for it in st.session_state.adj_cart:
                        over_qty   = it["qty"] if "Over" in it["type"] else 0
                        spoil_qty  = it["qty"] if "Spoilage" in it["type"] else 0
                        rows.append([ts, adj_date.strftime("%Y-%m-%d"),
                                     adj_date.strftime("%b %Y").upper(),
                                     it["item"], adj_staff.strip(), "ADJUSTMENT", ref,
                                     0, over_qty, 0, 0, 0, 0, 0, spoil_qty, it["notes"]])
                    log_ws.append_rows(rows)
                    invalidate_cache()
                    n = len(st.session_state.adj_cart)
                    st.session_state.adj_cart = []
                    st.success(f"✅ Adjustment **{ref}** submitted! {n} item(s) updated.")
                    st.balloons()
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MONTHLY REPORT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📆 Generate Report":
    st.markdown('<div class="main-header"><h1>📆 Generate Report</h1><p>Live inventory snapshot — generate anytime</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    log_df   = load_log(SPREADSHEET_ID)

    if items_df.empty:
        st.warning("No items yet."); st.stop()

    # Get available months from log
    if not log_df.empty and "MONTH" in log_df.columns:
        months = sorted(log_df["MONTH"].dropna().unique().tolist(), reverse=True)
    else:
        months = [date.today().strftime("%b %Y").upper()]

    c1, c2 = st.columns([2,3])
    with c1:
        sel_month = st.selectbox("Select Month", months)
    with c2:
        cat_filter = st.multiselect("Filter by Category", ["All"] + CATEGORIES, default=["All"])

    # Filter log for month
    month_log = log_df[log_df["MONTH"] == sel_month] if not log_df.empty else pd.DataFrame()

    # Build summary
    rows = []
    active = items_df[items_df["ACTIVE"]=="YES"]
    for _, item in active.iterrows():
        name = item["ITEM"]
        cost = num(item["UNIT COST"])
        beg  = num(item["BEGINNING_STOCKS"])
        cat  = item["CATEGORY"]

        if not month_log.empty:
            ilog = month_log[month_log["ITEM"]==name]
            add_in  = ilog["ADD_IN"].apply(num).sum()
            over    = ilog["OVER"].apply(num).sum()
            rest    = ilog["RESTAURANT"].apply(num).sum()
            banq    = ilog["BANQUET"].apply(num).sum()
            cafe    = ilog["CAFE"].apply(num).sum()
            bar     = ilog["BAR"].apply(num).sum()
            others  = ilog["OTHERS"].apply(num).sum()
            spoil   = ilog["SPOILAGE"].apply(num).sum()
        else:
            add_in = over = rest = banq = cafe = bar = others = spoil = 0

        ending = beg + add_in + over - rest - banq - cafe - bar - others - spoil
        worth  = ending * cost

        rows.append({
            "ITEM": name, "CATEGORY": cat, "UNIT": item["UNIT OF MEASURE"],
            "UNIT COST": cost, "BEGINNING": beg,
            "ADD'L/IN": add_in, "OVER": over,
            "RESTAURANT": rest, "BANQUET": banq, "CAFÉ": cafe, "BAR": bar,
            "OTHERS": others, "SPOILAGE": spoil,
            "ENDING": round(ending,4), "TOTAL WORTH": round(worth,4)
        })

    df = pd.DataFrame(rows)
    if "All" not in cat_filter and cat_filter:
        df = df[df["CATEGORY"].isin(cat_filter)]

    # Metrics
    total_worth = df["TOTAL WORTH"].sum()
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Month", sel_month)
    with c2: st.metric("Total Items", len(df))
    with c3: st.metric("Total Ending Worth", f"₱{total_worth:,.2f}")

    st.markdown("---")
    st.dataframe(df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: ITEM HISTORY
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔍 Item History":
    st.markdown('<div class="main-header"><h1>🔍 Item History</h1><p>Full transaction log per item or per reference number</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    log_df   = load_log(SPREADSHEET_ID)

    tab1, tab2 = st.tabs(["🔍 Search by Item", "🧾 Search by Reference (PO / Delivery / Adjustment)"])

    with tab1:
        if items_df.empty:
            st.info("No items yet.")
        else:
            search = st.text_input("Search Item", placeholder="Type item name...")
            active_names = items_df[items_df["ACTIVE"]=="YES"]["ITEM"].tolist()
            filtered_names = [i for i in active_names if search.lower() in i.lower()] if search else active_names
            sel = st.selectbox("Select Item", filtered_names, key="hist_item")

            if sel:
                info = items_df[items_df["ITEM"]==sel].iloc[0]
                st.markdown(f'<div class="info-box">📦 <strong>{sel}</strong> &nbsp;|&nbsp; {info["UNIT OF MEASURE"]} &nbsp;|&nbsp; ₱{num(info["UNIT COST"]):.4f}/unit &nbsp;|&nbsp; {info["CATEGORY"].upper()} &nbsp;|&nbsp; Current Beginning: <strong>{num(info["BEGINNING_STOCKS"]):,.2f}</strong></div>', unsafe_allow_html=True)

                if log_df.empty:
                    st.info("No transactions yet.")
                else:
                    ilog = log_df[log_df["ITEM"]==sel].copy()
                    if ilog.empty:
                        st.info("No transactions recorded for this item yet.")
                    else:
                        # Summary metrics
                        total_in    = ilog["ADD_IN"].apply(num).sum()
                        total_over  = ilog["OVER"].apply(num).sum()
                        total_out   = (ilog["RESTAURANT"].apply(num) + ilog["BANQUET"].apply(num) +
                                       ilog["CAFE"].apply(num) + ilog["BAR"].apply(num) +
                                       ilog["OTHERS"].apply(num)).sum()
                        total_spoil = ilog["SPOILAGE"].apply(num).sum()
                        beginning   = num(info["BEGINNING_STOCKS"])
                        ending_stock = beginning + total_in + total_over - total_out - total_spoil

                        c1,c2,c3,c4,c5 = st.columns(5)
                        with c1: st.metric("Total Incoming", f"{total_in:,.2f}")
                        with c2: st.metric("Total Released (PO)", f"{total_out:,.2f}")
                        with c3: st.metric("Total Spoilage", f"{total_spoil:,.2f}")
                        with c4: st.metric("Ending Stock", f"{ending_stock:,.2f}")
                        with c5: st.metric("Total Transactions", len(ilog))

                        st.markdown("---")

                        # Group by transaction type
                        for txn_type, icon, color in [
                            ("DELIVERY",   "📥", "#4A8ACC"),
                            ("PO",         "📋", "#8CAF7A"),
                            ("ADJUSTMENT", "🔧", "#C4A840"),
                            ("CARRYOVER",  "🔄", "#AA6ACC"),
                        ]:
                            tlog = ilog[ilog["TXN_TYPE"]==txn_type].sort_values("TIMESTAMP", ascending=False)
                            if tlog.empty:
                                continue

                            st.markdown(f'<div style="font-size:0.65rem;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:{color};margin:12px 0 6px 0;padding-bottom:4px;border-bottom:1px solid #1E2E1C;">{icon} {txn_type} — {len(tlog)} record(s)</div>', unsafe_allow_html=True)

                            for _, row in tlog.iterrows():
                                ref  = row.get("REF_NUMBER","")
                                date_str = row.get("DATE","")
                                staff = row.get("STAFF","")
                                parts = []
                                for col, label in [("ADD_IN","Incoming"),("OVER","Over (Adjustment)"),
                                                   ("RESTAURANT","Restaurant"),("BANQUET","Banquet"),
                                                   ("CAFE","Café"),("BAR","Bar"),
                                                   ("OTHERS","Others"),("SPOILAGE","Spoilage")]:
                                    v = num(row.get(col,0))
                                    if v: parts.append(f"<strong style='color:{color};'>{label}:</strong> {v:,.2f}")
                                _notes_val = str(row.get("NOTES", "")).strip()
                                notes_str = f'<div style="color:#5A7A52;font-size:0.78rem;margin-top:3px;">📝 {_notes_val}</div>' if _notes_val else ""
                                st.markdown(f'<div class="log-entry" style="border-left-color:{color};"><div style="display:flex;justify-content:space-between;align-items:center;"><span>📅 <strong>{date_str}</strong> &nbsp;·&nbsp; 👤 {staff}</span><span style="color:#3A5238;font-size:0.75rem;">{ref}</span></div><div style="margin-top:5px;">{" &nbsp;&nbsp; ".join(parts) if parts else "—"}</div>{notes_str}</div>', unsafe_allow_html=True)

    with tab2:
        ref_search = st.text_input("Search Reference #", placeholder="e.g. PO-20260628, DEL-..., ADJ-...")
        if ref_search and not log_df.empty:
            rlog = log_df[log_df["REF_NUMBER"].astype(str).str.contains(ref_search, case=False, na=False)]
            if rlog.empty:
                st.warning("No records found.")
            else:
                for ref in rlog["REF_NUMBER"].unique():
                    rr = rlog[rlog["REF_NUMBER"]==ref]
                    first = rr.iloc[0]
                    st.markdown(f'<div class="info-box">🧾 <strong>{ref}</strong> &nbsp;|&nbsp; {first.get("TXN_TYPE","")} &nbsp;|&nbsp; 📅 {first.get("DATE","")} &nbsp;|&nbsp; 👤 {first.get("STAFF","")} &nbsp;|&nbsp; {len(rr)} item(s)</div>', unsafe_allow_html=True)
                    for _, row in rr.iterrows():
                        parts = []
                        for col, label in [("ADD_IN","Incoming"),("OVER","Over"),
                                           ("RESTAURANT","Restaurant"),("BANQUET","Banquet"),
                                           ("CAFE","Café"),("BAR","Bar"),("OTHERS","Others"),("SPOILAGE","Spoilage")]:
                            v = num(row.get(col,0))
                            if v: parts.append(f"{label}: <strong>{v:,.2f}</strong>")
                        st.markdown(f'<div class="log-entry">📦 <strong>{row.get("ITEM","")}</strong><div style="margin-top:3px;">{" &nbsp;·&nbsp; ".join(parts) if parts else "—"}</div></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: ITEMS MASTER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📦 Items Master":
    st.markdown('<div class="main-header"><h1>📦 Items Master</h1><p>Manage your ingredient list</p></div>', unsafe_allow_html=True)

    items_df = load_items(SPREADSHEET_ID)
    tab1, tab2, tab3 = st.tabs(["➕ Add New Item", "✏️ Edit / Deactivate", "📋 View All"])

    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            new_name = st.text_input("Item Name *")
            new_unit = st.text_input("Unit of Measure *", placeholder="gram / ml / piece / bottle")
        with c2:
            new_cost = st.number_input("Unit Cost (₱) *", min_value=0.0, step=0.0001, format="%.4f")
            new_cat  = st.selectbox("Category *", CATEGORIES)
        new_begin = st.number_input("Beginning Stock", min_value=0.0, step=0.01, format="%.2f")

        if st.button("➕ Add Item", type="primary"):
            if not new_name.strip() or not new_unit.strip():
                st.error("Name and unit are required.")
            elif not items_df.empty and new_name.strip() in items_df["ITEM"].values:
                st.error("Item already exists.")
            else:
                ws = ensure_sheet(ss, ITEMS_SHEET, ITEMS_HEADERS)
                ws.append_row([new_name.strip(), new_cost, new_unit.strip(), new_cat, new_begin, "YES"])
                invalidate_cache()
                st.success(f"✅ {new_name} added!"); st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        if items_df.empty:
            st.info("No items yet.")
        else:
            search = st.text_input("🔍 Search", key="edit_search")
            fdf = items_df[items_df["ITEM"].str.contains(search, case=False, na=False)] if search else items_df
            sel = st.selectbox("Select Item", fdf["ITEM"].tolist())
            if sel:
                row = items_df[items_df["ITEM"]==sel].iloc[0]
                ridx = items_df[items_df["ITEM"]==sel].index[0] + 2
                st.markdown('<div class="card">', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    e_cost  = st.number_input("Unit Cost", value=num(row["UNIT COST"]), step=0.0001, format="%.4f", key="e_cost")
                    e_unit  = st.text_input("Unit of Measure", value=str(row["UNIT OF MEASURE"]), key="e_unit")
                    e_begin = st.number_input("Beginning Stock", value=num(row["BEGINNING_STOCKS"]), step=0.01, format="%.2f", key="e_begin")
                with c2:
                    e_cat    = st.selectbox("Category", CATEGORIES, index=CATEGORIES.index(row["CATEGORY"]) if row["CATEGORY"] in CATEGORIES else 0, key="e_cat")
                    e_active = st.selectbox("Status", ["YES","NO"], index=0 if row["ACTIVE"]=="YES" else 1, key="e_active")
                if st.button("💾 Save Changes", type="primary"):
                    ws = ss.worksheet(ITEMS_SHEET)
                    ws.update_cell(ridx, 2, e_cost)
                    ws.update_cell(ridx, 3, e_unit)
                    ws.update_cell(ridx, 4, e_cat)
                    ws.update_cell(ridx, 5, e_begin)
                    ws.update_cell(ridx, 6, e_active)
                    invalidate_cache()
                    st.success(f"✅ {sel} updated!"); st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    with tab3:
        if items_df.empty:
            st.info("No items yet.")
        else:
            show_all = st.checkbox("Show inactive items")
            df = items_df if show_all else items_df[items_df["ACTIVE"]=="YES"]
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption(f"{len(df)} items")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: EXPORT TO EXCEL
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⬇️ Export to Excel":
    st.markdown('<div class="main-header"><h1>⬇️ Export to Excel</h1><p>Generate inventory report for any date range</p></div>', unsafe_allow_html=True)

    import calendar

    items_df = load_items(SPREADSHEET_ID)
    log_df   = load_log(SPREADSHEET_ID)

    if items_df.empty:
        st.warning("No items yet. Go to Setup first."); st.stop()

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Select Date Range</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        from_date = st.date_input("📅 From Date", value=date.today().replace(day=1), key="exp_from")
    with c2:
        to_date = st.date_input("📅 To Date", value=date.today(), key="exp_to")

    if from_date > to_date:
        st.error("From date must be before To date.")
        st.stop()

    # Show how many days and transactions in range
    from_str = from_date.strftime("%Y-%m-%d")
    to_str   = to_date.strftime("%Y-%m-%d")

    if not log_df.empty and "DATE" in log_df.columns:
        range_log = log_df[(log_df["DATE"] >= from_str) & (log_df["DATE"] <= to_str)]
        st.markdown(f'<div class="info-box">📊 <strong>{len(range_log)}</strong> transactions found between <strong>{from_date.strftime("%b %d, %Y")}</strong> and <strong>{to_date.strftime("%b %d, %Y")}</strong></div>', unsafe_allow_html=True)
    else:
        range_log = pd.DataFrame()
        st.markdown('<div class="info-box">No transactions found in this range.</div>', unsafe_allow_html=True)

    if st.button("📥 Generate Excel File", type="primary", use_container_width=True):

        # Build list of all days in range
        from datetime import timedelta
        all_days = []
        cur = from_date
        while cur <= to_date:
            all_days.append(cur.strftime("%Y-%m-%d"))
            cur += timedelta(days=1)

        # File label
        if from_date.month == to_date.month and from_date.year == to_date.year:
            file_label = from_date.strftime("%B_%Y")
            title_label = from_date.strftime("%B %Y").upper()
        else:
            file_label = f"{from_date.strftime('%b_%Y')}_to_{to_date.strftime('%b_%Y')}"
            title_label = f"{from_date.strftime('%b %d').upper()} TO {to_date.strftime('%b %d, %Y').upper()}"

        active = items_df[items_df["ACTIVE"] == "YES"].copy().reset_index(drop=True)

        wb = openpyxl.Workbook()

        # ── Styles ──────────────────────────────────────────────────────────
        hdr_fill   = PatternFill("solid", fgColor="1A2E1A")
        title_fill = PatternFill("solid", fgColor="0A1208")
        worth_fill = PatternFill("solid", fgColor="0E1C0E")
        hdr_font   = Font(bold=True, color="A8C896", size=9)
        title_font = Font(bold=True, color="C8DCC0", size=12)
        worth_font = Font(bold=True, color="8CAF7A", size=9)
        thin       = Side(style="thin", color="2A3828")
        border     = Border(left=thin, right=thin, top=thin, bottom=thin)
        num_fmt4   = "#,##0.0000"

        cat_colors = {
            "beverage":"D6E8F5","beef":"F5D6D6","chicken":"FFF3D6",
            "seafood":"D6F0F5","fresh":"D6F5E0","dry":"FFF8D6",
            "wet":"EDD6F5","rtc":"D6E8FF","pork":"F5D6D6",
            "meat":"F0F0F0","frozen":"D6EAF8","dessert":"FFF0D6"
        }

        MAIN_COLS  = ["ITEM","BEGINNING STOCKS","UNIT COST","UNIT OF MEASURE","CATEGORY",
                      "ADD'L / IN","OVER","RESTAURANT","BANQUET","CAFÉ","BAR","OTHERS",
                      "SPOILAGE","ENDING STOCKS","TOTAL WORTH OF STOCKS"]
        WORTH_COLS = ["WORTH OF BEGINNING","WORTH OF ADD'L/IN","WORTH OF OVER",
                      "WORTH OF RESTAURANT","WORTH OF BANQUET","WORTH OF CAFÉ",
                      "WORTH OF BAR","WORTH OF OTHERS","WORTH OF SPOILAGE","WORTH OF ENDING"]

        def col_letter(n): return get_column_letter(n)

        # Running stock tracker — beginning of range = current BEGINNING_STOCKS
        item_running = {row["ITEM"]: num(row["BEGINNING_STOCKS"]) for _, row in active.iterrows()}

        def write_day_sheet(ws_out, day_str, day_label, tab_log, item_running):
            ws_out.merge_cells("A1:O1")
            tc = ws_out["A1"]
            tc.value = f"{day_label} — SERVANDO MAIN WAREHOUSE INVENTORY"
            tc.font = title_font; tc.fill = title_fill
            tc.alignment = Alignment(horizontal="center", vertical="center")
            ws_out.row_dimensions[1].height = 26

            for ci, col in enumerate(MAIN_COLS, 1):
                c = ws_out.cell(row=2, column=ci, value=col)
                c.fill = hdr_fill; c.font = hdr_font
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = border
            for ci, col in enumerate(WORTH_COLS, 17):
                c = ws_out.cell(row=2, column=ci, value=col)
                c.fill = worth_fill; c.font = worth_font
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = border
            ws_out.row_dimensions[2].height = 38

            for ri, (_, item) in enumerate(active.iterrows(), 3):
                name = item["ITEM"]
                cost = num(item["UNIT COST"])
                cat  = str(item["CATEGORY"]).lower()
                beg  = item_running.get(name, 0.0)

                if not tab_log.empty:
                    ilog   = tab_log[tab_log["ITEM"] == name]
                    add_in = ilog["ADD_IN"].apply(num).sum()
                    over   = ilog["OVER"].apply(num).sum()
                    rest   = ilog["RESTAURANT"].apply(num).sum()
                    banq   = ilog["BANQUET"].apply(num).sum()
                    cafe   = ilog["CAFE"].apply(num).sum()
                    bar    = ilog["BAR"].apply(num).sum()
                    others = ilog["OTHERS"].apply(num).sum()
                    spoil  = ilog["SPOILAGE"].apply(num).sum()
                else:
                    add_in=over=rest=banq=cafe=bar=others=spoil=0.0

                ending = beg + add_in + over - rest - banq - cafe - bar - others - spoil
                worth  = ending * cost
                cfill  = PatternFill("solid", fgColor=cat_colors.get(cat, "FFFFFF"))

                main_vals = [name, round(beg,4), round(cost,6), item["UNIT OF MEASURE"],
                             item["CATEGORY"], round(add_in,4), round(over,4),
                             round(rest,4), round(banq,4), round(cafe,4), round(bar,4),
                             round(others,4), round(spoil,4), round(ending,4), round(worth,4)]
                worth_vals = [beg*cost, add_in*cost, over*cost, rest*cost, banq*cost,
                              cafe*cost, bar*cost, others*cost, spoil*cost, worth]

                for ci, val in enumerate(main_vals, 1):
                    cell = ws_out.cell(row=ri, column=ci, value=val)
                    cell.fill = cfill; cell.border = border
                    if ci in [2,3,6,7,8,9,10,11,12,13,14,15]:
                        cell.number_format = num_fmt4
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(vertical="center")

                for ci, val in enumerate(worth_vals, 17):
                    cell = ws_out.cell(row=ri, column=ci, value=round(val,4))
                    cell.fill = worth_fill
                    cell.font = Font(color="8CAF7A", size=9)
                    cell.number_format = num_fmt4; cell.border = border
                    cell.alignment = Alignment(horizontal="right", vertical="center")

            # Totals row
            total_row = len(active) + 3
            ws_out.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True, color="FFFFFF")
            ws_out.cell(row=total_row, column=1).fill = hdr_fill
            ws_out.cell(row=total_row, column=1).border = border
            for ci in [2,6,7,8,9,10,11,12,13,14,15]:
                col_l = col_letter(ci)
                cell = ws_out.cell(row=total_row, column=ci,
                    value=f"=SUM({col_l}3:{col_l}{total_row-1})")
                cell.fill = hdr_fill; cell.font = Font(bold=True, color="A8C896", size=9)
                cell.number_format = num_fmt4; cell.border = border
                cell.alignment = Alignment(horizontal="right")
            for ci in range(17, 17+len(WORTH_COLS)):
                col_l = col_letter(ci)
                cell = ws_out.cell(row=total_row, column=ci,
                    value=f"=SUM({col_l}3:{col_l}{total_row-1})")
                cell.fill = worth_fill; cell.font = Font(bold=True, color="8CAF7A", size=9)
                cell.number_format = num_fmt4; cell.border = border
                cell.alignment = Alignment(horizontal="right")

            col_widths = [32,13,11,14,11,11,10,12,12,10,10,10,12,14,17,3,
                          14,13,12,14,14,12,12,12,14,14]
            for ci, w in enumerate(col_widths, 1):
                ws_out.column_dimensions[col_letter(ci)].width = w
            ws_out.freeze_panes = "A3"

            # Update running stock
            for _, item in active.iterrows():
                name = item["ITEM"]
                beg  = item_running.get(name, 0.0)
                if not tab_log.empty:
                    ilog   = tab_log[tab_log["ITEM"] == name]
                    add_in = ilog["ADD_IN"].apply(num).sum()
                    over   = ilog["OVER"].apply(num).sum()
                    rest   = ilog["RESTAURANT"].apply(num).sum()
                    banq   = ilog["BANQUET"].apply(num).sum()
                    cafe   = ilog["CAFE"].apply(num).sum()
                    bar    = ilog["BAR"].apply(num).sum()
                    others = ilog["OTHERS"].apply(num).sum()
                    spoil  = ilog["SPOILAGE"].apply(num).sum()
                else:
                    add_in=over=rest=banq=cafe=bar=others=spoil=0.0
                item_running[name] = beg + add_in + over - rest - banq - cafe - bar - others - spoil
            return item_running

        # ── Write one tab per day ────────────────────────────────────────
        for day_str in all_days:
            try:
                d = datetime.strptime(day_str, "%Y-%m-%d")
                tab_name  = str(d.day)
                day_label = d.strftime("%B %d, %Y").upper()
            except:
                tab_name  = day_str[-2:].lstrip("0") or "1"
                day_label = day_str

            tab_log = range_log[range_log["DATE"] == day_str] if not range_log.empty else pd.DataFrame()
            ws_day  = wb.create_sheet(title=tab_name)
            item_running = write_day_sheet(ws_day, day_str, day_label, tab_log, item_running)

        # ── SUMMARY tab ──────────────────────────────────────────────────
        ws_sum = wb.create_sheet(title="SUMMARY")
        ws_sum.merge_cells("A1:O1")
        tc = ws_sum["A1"]
        tc.value = f"SUMMARY — {title_label} — SERVANDO MAIN WAREHOUSE INVENTORY"
        tc.font = title_font; tc.fill = title_fill
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws_sum.row_dimensions[1].height = 26

        for ci, col in enumerate(MAIN_COLS, 1):
            c = ws_sum.cell(row=2, column=ci, value=col)
            c.fill = hdr_fill; c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
        for ci, col in enumerate(WORTH_COLS, 17):
            c = ws_sum.cell(row=2, column=ci, value=col)
            c.fill = worth_fill; c.font = worth_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
        ws_sum.row_dimensions[2].height = 38

        for ri, (_, item) in enumerate(active.iterrows(), 3):
            name = item["ITEM"]
            cost = num(item["UNIT COST"])
            cat  = str(item["CATEGORY"]).lower()
            beg  = num(item["BEGINNING_STOCKS"])

            if not range_log.empty:
                ilog   = range_log[range_log["ITEM"] == name]
                add_in = ilog["ADD_IN"].apply(num).sum()
                over   = ilog["OVER"].apply(num).sum()
                rest   = ilog["RESTAURANT"].apply(num).sum()
                banq   = ilog["BANQUET"].apply(num).sum()
                cafe   = ilog["CAFE"].apply(num).sum()
                bar    = ilog["BAR"].apply(num).sum()
                others = ilog["OTHERS"].apply(num).sum()
                spoil  = ilog["SPOILAGE"].apply(num).sum()
            else:
                add_in=over=rest=banq=cafe=bar=others=spoil=0.0

            ending = beg + add_in + over - rest - banq - cafe - bar - others - spoil
            worth  = ending * cost
            cfill  = PatternFill("solid", fgColor=cat_colors.get(cat, "FFFFFF"))

            main_vals = [name, round(beg,4), round(cost,6), item["UNIT OF MEASURE"],
                         item["CATEGORY"], round(add_in,4), round(over,4),
                         round(rest,4), round(banq,4), round(cafe,4), round(bar,4),
                         round(others,4), round(spoil,4), round(ending,4), round(worth,4)]
            worth_vals = [beg*cost, add_in*cost, over*cost, rest*cost, banq*cost,
                          cafe*cost, bar*cost, others*cost, spoil*cost, worth]

            for ci, val in enumerate(main_vals, 1):
                cell = ws_sum.cell(row=ri, column=ci, value=val)
                cell.fill = cfill; cell.border = border
                if ci in [2,3,6,7,8,9,10,11,12,13,14,15]:
                    cell.number_format = num_fmt4
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(vertical="center")
            for ci, val in enumerate(worth_vals, 17):
                cell = ws_sum.cell(row=ri, column=ci, value=round(val,4))
                cell.fill = worth_fill
                cell.font = Font(color="8CAF7A", size=9)
                cell.number_format = num_fmt4; cell.border = border
                cell.alignment = Alignment(horizontal="right", vertical="center")

        total_row = len(active) + 3
        ws_sum.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True, color="FFFFFF")
        ws_sum.cell(row=total_row, column=1).fill = hdr_fill
        ws_sum.cell(row=total_row, column=1).border = border
        for ci in [2,6,7,8,9,10,11,12,13,14,15]:
            col_l = col_letter(ci)
            cell = ws_sum.cell(row=total_row, column=ci, value=f"=SUM({col_l}3:{col_l}{total_row-1})")
            cell.fill = hdr_fill; cell.font = Font(bold=True, color="A8C896", size=9)
            cell.number_format = num_fmt4; cell.border = border
            cell.alignment = Alignment(horizontal="right")
        for ci in range(17, 17+len(WORTH_COLS)):
            col_l = col_letter(ci)
            cell = ws_sum.cell(row=total_row, column=ci, value=f"=SUM({col_l}3:{col_l}{total_row-1})")
            cell.fill = worth_fill; cell.font = Font(bold=True, color="8CAF7A", size=9)
            cell.number_format = num_fmt4; cell.border = border
            cell.alignment = Alignment(horizontal="right")

        col_widths = [32,13,11,14,11,11,10,12,12,10,10,10,12,14,17,3,
                      14,13,12,14,14,12,12,12,14,14]
        for ci, w in enumerate(col_widths, 1):
            ws_sum.column_dimensions[col_letter(ci)].width = w
        ws_sum.freeze_panes = "A3"

        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        buf = BytesIO()
        wb.save(buf); buf.seek(0)

        st.download_button(
            label=f"📥 Download {file_label.replace('_', ' ')}.xlsx",
            data=buf,
            file_name=f"{file_label}_-_Servando_Main_Warehouse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"✅ Done! {len(all_days)} daily tabs + SUMMARY. File: {file_label} - Servando Main Warehouse.xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SETUP
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Setup":
    st.markdown('<div class="main-header"><h1>⚙️ Setup</h1><p>Initialize system · Import items</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Import 557 Items from May 2026 Inventory</div>', unsafe_allow_html=True)
    st.write("This imports all your existing ingredients into the new system. Run once. Skips items that already exist.")

    if st.button("🚀 Import Items", type="primary"):
        ws = ensure_sheet(ss, ITEMS_SHEET, ITEMS_HEADERS)
        existing = {r["ITEM"] for r in ws.get_all_records()}
        new_rows = []
        for item in INITIAL_ITEMS:
            name, beginning, cost, uom, cat = item
            if name not in existing:
                new_rows.append([name, round(cost,6), uom, cat, round(beginning,4), "YES"])

        if new_rows:
            # Batch in groups of 100 to avoid rate limits
            for i in range(0, len(new_rows), 100):
                ws.append_rows(new_rows[i:i+100])
            invalidate_cache()
            st.success(f"✅ Imported {len(new_rows)} items!")
        else:
            st.info("All items already exist.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Current Status</div>', unsafe_allow_html=True)
    items_df = load_items(SPREADSHEET_ID)
    log_df   = load_log(SPREADSHEET_ID)
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Items in Master", len(items_df))
    with c2: st.metric("Total Transactions", len(log_df))
    with c3: st.metric("Active Items", len(items_df[items_df["ACTIVE"]=="YES"]) if not items_df.empty else 0)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="section-title" style="color:#CC6A6A;">⚠️ Danger Zone — Reset / Clear Data</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:#CC6A6A;font-size:0.82rem;margin-bottom:1rem;">These actions are permanent and cannot be undone. Type the confirmation word before proceeding.</div>', unsafe_allow_html=True)

    # Clear Transactions only
    st.markdown('<div class="card" style="border-color:#3A1A1A;">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Clear All Transactions (Keep Items)</div>', unsafe_allow_html=True)
    st.write("Wipes the entire transaction log (deliveries, POs, adjustments). Your item list stays intact.")
    confirm1 = st.text_input("Type **CLEAR TRANSACTIONS** to confirm", key="confirm_txn")
    if st.button("🗑️ Clear Transactions", key="btn_clear_txn"):
        if confirm1.strip().upper() == "CLEAR TRANSACTIONS":
            ws = ensure_sheet(ss, LOG_SHEET, LOG_HEADERS)
            ws.clear()
            ws.append_row(LOG_HEADERS)
            invalidate_cache()
            st.success("✅ All transactions cleared. Items list is untouched.")
        else:
            st.error("❌ Confirmation text doesn't match. Nothing was deleted.")
    st.markdown('</div>', unsafe_allow_html=True)

    # Clear Items only
    st.markdown('<div class="card" style="border-color:#3A1A1A;">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Clear All Items (Keep Transactions)</div>', unsafe_allow_html=True)
    st.write("Wipes the entire items master list. Transaction history stays intact.")
    confirm2 = st.text_input("Type **CLEAR ITEMS** to confirm", key="confirm_items")
    if st.button("🗑️ Clear Items", key="btn_clear_items"):
        if confirm2.strip().upper() == "CLEAR ITEMS":
            ws = ensure_sheet(ss, ITEMS_SHEET, ITEMS_HEADERS)
            ws.clear()
            ws.append_row(ITEMS_HEADERS)
            invalidate_cache()
            st.success("✅ All items cleared. Transaction log is untouched.")
        else:
            st.error("❌ Confirmation text doesn't match. Nothing was deleted.")
    st.markdown('</div>', unsafe_allow_html=True)

    # Full Reset
    st.markdown('<div class="card" style="border-color:#5A1A1A;">', unsafe_allow_html=True)
    st.markdown('<div class="section-title" style="color:#CC6A6A;">Full Reset — Clear Everything</div>', unsafe_allow_html=True)
    st.write("Wipes **both** the items list and all transactions. Complete fresh start.")
    confirm3 = st.text_input("Type **FULL RESET** to confirm", key="confirm_full")
    if st.button("💥 Full Reset", key="btn_full_reset"):
        if confirm3.strip().upper() == "FULL RESET":
            for sheet_name, headers in [(ITEMS_SHEET, ITEMS_HEADERS), (LOG_SHEET, LOG_HEADERS)]:
                ws = ensure_sheet(ss, sheet_name, headers)
                ws.clear()
                ws.append_row(headers)
            invalidate_cache()
            st.success("✅ Full reset complete. Everything has been cleared.")
        else:
            st.error("❌ Confirmation text doesn't match. Nothing was deleted.")
    st.markdown('</div>', unsafe_allow_html=True)