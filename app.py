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


def build_doc_xlsx(doc_type, ref, date_str, staff_str, meta3_label, meta3_val, rows_data, logo_b64_str):
    """
    Build a PO or Delivery xlsx matching the exact template format.
    rows_data: list of (num, item_name, unit, qty, amount)
    Returns BytesIO buffer.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.drawing.image import Image as XLImage
    from io import BytesIO
    import base64

    wb = Workbook()
    ws = wb.active
    ws.title = doc_type

    # ── fonts ──
    f14b  = Font(bold=True, name="Arial", size=14)
    f8g   = Font(name="Arial", size=8,  color="888888")
    f12g  = Font(bold=True, name="Arial", size=12, color="888888")
    f12b  = Font(bold=True, name="Arial", size=12)
    f9w   = Font(bold=True, name="Arial", size=9,  color="FFFFFF")
    f11   = Font(name="Arial", size=11)
    f11b  = Font(bold=True, name="Arial", size=11)
    hdr_fill = PatternFill("solid", fgColor="1A2E1A")
    tot_fill = PatternFill("solid", fgColor="E8F0E8")
    ctr  = Alignment(horizontal="center")
    rgt  = Alignment(horizontal="right")
    vctr = Alignment(vertical="center")

    # ── col widths / row heights ──
    ws.column_dimensions["A"].width = 5.0
    ws.column_dimensions["B"].width = 38.0
    ws.column_dimensions["C"].width = 10.90625
    ws.column_dimensions["D"].width = 12.0
    ws.column_dimensions["E"].width = 16.0

    ws.row_dimensions[1].height = 22.0
    ws.row_dimensions[2].height = 14.0
    ws.row_dimensions[3].height = 11.0
    ws.row_dimensions[4].height = 18.0
    ws.row_dimensions[5].height = 16.0
    ws.row_dimensions[6].height = 8.0
    ws.row_dimensions[7].height = 18.0

    # ── logo ──
    try:
        logo_buf = BytesIO(base64.b64decode(logo_b64_str))
        xl_logo = XLImage(logo_buf)
        xl_logo.width = 220
        xl_logo.height = 58
        xl_logo.anchor = "D1"
        ws.add_image(xl_logo)
    except Exception:
        pass

    # ── row 1: title ──
    ws.merge_cells("A1:E1")
    ws["A1"] = doc_type
    ws["A1"].font = f14b
    ws["A1"].alignment = vctr

    # ── row 2: reference ──
    ws.merge_cells("A2:E2")
    ws["A2"] = f"Reference: {ref}"
    ws["A2"].font = f8g

    # ── rows 4–5: meta labels + values ──
    # meta1=Prepared By/Received By col B, meta2=Date col C, meta3=Dept/Items col E
    meta1_label = "Prepared By" if doc_type == "PURCHASE ORDER" else "Received By"
    ws["B4"] = meta1_label;  ws["B4"].font = f12g
    ws["C4"] = "Date";       ws["C4"].font = f12g
    ws["E4"] = meta3_label;  ws["E4"].font = f12g
    ws["B5"] = staff_str;    ws["B5"].font = f12b
    ws["C5"] = date_str;     ws["C5"].font = f12b
    ws["E5"] = meta3_val;    ws["E5"].font = f12b

    # ── row 7: table header ──
    for ci, (hdr, al) in enumerate([
        ("#", ctr), ("Item", None), ("Unit", ctr), ("Qty", ctr), ("Amount", ctr)
    ], 1):
        c = ws.cell(row=7, column=ci, value=hdr)
        c.font = f9w
        c.fill = hdr_fill
        c.alignment = al or Alignment()

    # ── rows 8–40: data rows (30 slots) ──
    for ri in range(30):
        row_num = 8 + ri
        ws.row_dimensions[row_num].height = 18.0 if ri == 0 else None
        if ri < len(rows_data):
            num_i, item_n, unit, qty, amt = rows_data[ri]
            qty_val = int(qty) if qty == int(qty) else round(qty, 2)
            vals = [num_i, item_n, unit, qty_val, amt]
            aligns = [ctr, None, ctr, ctr, rgt]
            fmts   = [None, None, None, "#,##0", "₱#,##0.00"]
            for ci, (v, al, fmt) in enumerate(zip(vals, aligns, fmts), 1):
                c = ws.cell(row=row_num, column=ci, value=v)
                c.font = f11
                if al: c.alignment = al
                if fmt: c.number_format = fmt

    # ── row 41: total (always fixed row) ──
    ws["A41"].font = f11b
    ws["B41"].font = f11b
    ws["C41"].font = f11b
    ws["D41"].font = f11b
    ws.merge_cells("A41:D41")
    ws["A41"] = "TOTAL"
    ws["A41"].alignment = rgt
    total = sum(r[4] for r in rows_data)
    ws["E41"] = total
    ws["E41"].font = f11b
    ws["E41"].fill = tot_fill
    ws["E41"].alignment = rgt
    ws["E41"].number_format = "₱#,##0.00"

    # ── row 43: signature labels ──
    if doc_type == "PURCHASE ORDER":
        sigs = [("A43","Prepared By"), ("C43","Approved By"), ("E43","Received By")]
    else:
        sigs = [("A43","Received By"), ("C43","Checked By"),  ("E43","Noted By")]
    for coord, label in sigs:
        ws[coord] = label
        ws[coord].font = f8g

    # ── print setup ──
    ws.print_area = "A1:E43"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

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

    # ── Past Deliveries ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">📂 Past Deliveries</div>', unsafe_allow_html=True)
    log_df = load_log(SPREADSHEET_ID)
    if not log_df.empty:
        del_log = log_df[log_df["TXN_TYPE"] == "DELIVERY"]
        if del_log.empty:
            st.info("No past deliveries yet.")
        else:
            past_refs = sorted(del_log["REF_NUMBER"].unique().tolist(), reverse=True)
            sel_ref = st.selectbox("Select Delivery Reference", past_refs, key="past_del_ref")
            if sel_ref:
                ref_rows = del_log[del_log["REF_NUMBER"] == sel_ref]
                first = ref_rows.iloc[0]
                total_val = sum(num(r["ADD_IN"]) * num(items_df[items_df["ITEM"]==r["ITEM"]]["UNIT COST"].values[0]) if r["ITEM"] in items_df["ITEM"].values else 0 for _, r in ref_rows.iterrows())
                st.markdown(f'<div class="info-box">📥 <strong>{sel_ref}</strong> &nbsp;|&nbsp; 📅 {first.get("DATE","")} &nbsp;|&nbsp; 👤 {first.get("STAFF","")} &nbsp;|&nbsp; {len(ref_rows)} item(s)</div>', unsafe_allow_html=True)

                table_rows = ""
                for row_idx, (_, r) in enumerate(ref_rows.iterrows(), 1):
                    qty = num(r["ADD_IN"])
                    unit_cost = num(items_df[items_df["ITEM"]==r["ITEM"]]["UNIT COST"].values[0]) if r["ITEM"] in items_df["ITEM"].values else 0
                    uom = items_df[items_df["ITEM"]==r["ITEM"]]["UNIT OF MEASURE"].values[0] if r["ITEM"] in items_df["ITEM"].values else ""
                    val = qty * unit_cost
                    st.markdown(f'<div class="po-item-row"><strong>{r["ITEM"]}</strong> &nbsp;|&nbsp; <span style="color:#8CAF7A;">{qty:,.2f} {uom}</span> &nbsp;|&nbsp; ₱{unit_cost:.4f}/unit &nbsp;|&nbsp; ₱{val:,.2f}</div>', unsafe_allow_html=True)
                    table_rows += f"<tr><td>{row_idx}</td><td>{r['ITEM']}</td><td>{uom}</td><td class='right'>{qty:,.2f}</td><td class='right'>&#8369;{unit_cost:.4f}</td><td class='right'>&#8369;{val:,.2f}</td></tr>"

                st.markdown(f'<div class="info-box" style="text-align:right;">Total Value: <strong>₱{total_val:,.2f}</strong></div>', unsafe_allow_html=True)

                del_rows = []
                for ri, (_, r) in enumerate(ref_rows.iterrows(), 1):
                    qty = num(r["ADD_IN"])
                    unit_cost = num(items_df[items_df["ITEM"]==r["ITEM"]]["UNIT COST"].values[0]) if r["ITEM"] in items_df["ITEM"].values else 0
                    uom = items_df[items_df["ITEM"]==r["ITEM"]]["UNIT OF MEASURE"].values[0] if r["ITEM"] in items_df["ITEM"].values else ""
                    del_rows.append((ri, r["ITEM"], uom, float(qty), qty * unit_cost))
                buf = build_doc_xlsx("INCOMING DELIVERY", sel_ref, first.get("DATE",""), first.get("STAFF",""), "Items", str(len(ref_rows)), del_rows, "/9j/4QA4RXhpZgAASUkqAAgAAAABAGmHBAABAAAAHAAAAAAAAAAAAAEAAJAHAAQAAAAwMjEwAAAAAAAA/+AAEEpGSUYAAQEAAAEAAQAA/+IB2ElDQ19QUk9GSUxFAAEBAAAByAAAAAAEMAAAbW50clJHQiBYWVogB+AAAQABAAAAAAAAYWNzcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAPbWAAEAAAAA0y0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJZGVzYwAAAPAAAAAkclhZWgAAARQAAAAUZ1hZWgAAASgAAAAUYlhZWgAAATwAAAAUd3RwdAAAAVAAAAAUclRSQwAAAWQAAAAoZ1RSQwAAAWQAAAAoYlRSQwAAAWQAAAAoY3BydAAAAYwAAAA8bWx1YwAAAAAAAAABAAAADGVuVVMAAAAIAAAAHABzAFIARwBCWFlaIAAAAAAAAG+iAAA49QAAA5BYWVogAAAAAAAAYpkAALeFAAAY2lhZWiAAAAAAAAAkoAAAD4QAALbPWFlaIAAAAAAAAPbWAAEAAAAA0y1wYXJhAAAAAAAEAAAAAmZmAADypwAADVkAABPQAAAKWwAAAAAAAAAAbWx1YwAAAAAAAAABAAAADGVuVVMAAAAgAAAAHABHAG8AbwBnAGwAZQAgAEkAbgBjAC4AIAAyADAAMQA2/9sAQwADAgICAgIDAgICAwMDAwQGBAQEBAQIBgYFBgkICgoJCAkJCgwPDAoLDgsJCQ0RDQ4PEBAREAoMEhMSEBMPEBAQ/9sAQwEDAwMEAwQIBAQIEAsJCxAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ/8AAEQgBkwcCAwEiAAIRAQMRAf/EAB0AAQACAgMBAQAAAAAAAAAAAAAICQEHAgUGAwT/xABnEAABAwIDBAQFCwwOBgcIAwEBAAIDBAUGBxEIEiExE0FRYQkUInGBFRgyN0J1kaGxs9EZI1JWYnJzgpKTlLIWFyQzNDU2OFNXdHa0wSVDVaKj4WSDwsPS8PEmJ0RFRlRjhChHpGX/xAAZAQEAAwEBAAAAAAAAAAAAAAAAAQIDBAX/xAAsEQEAAgIBAwQDAQEBAAMAAwAAAQIDERIEITETFDJRIjNBYUJSI3GBJDSR/9oADAMBAAIRAxEAPwC1NERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERARY1TTvUROxlFhE3AyiIm4BFj4UTlAyiIpBERAREQEREBERAREQEREBERAREQERYQZRflra+lt8DqqsqoaeFg8qSV4a0ek8FqDHm1vkrgPfhqMTtudWznTW5vTP8ASeAHwqlr1r5lG26UUJ8S+EZgjc6PCuAjJp7F9bPpr+K3T5VrW97fWddxc71MZabYw8hHTb5H4ztVhbq8df6jlCyJFVtNtn7QEr944yaw/YtpmNHyKRmydtZ4pzJxOMvcfMhnrZYny0lbEwML91pcWOA4chzU06ql7cSLbS9WD59EKr+2odprODDWbl3wnhq+zWe22oxRQthaAZdYw4vLjz4kjT7la5ckYq8pTM6WAhwPJZVYFh22c+bLM18uJILjEOcdXTNdr6dNVuLCXhFXno4saYIb91NQTHQ/iu1WNerpZWLxKbqLTGCNrbJTHDWMpsVR22qd/wDD3AdC70HkfhW2bdeLddoPGrXcKashPJ8Eoe0+kLeuStvErbfuRY4r81fWQ2+inraqdsMUEbnve9wAaAOZJV0vt0g9j0g15ceeq+iqSxfnRjy75m3DFVBjC5DW4yOpSypcGMj3/JaBrpporWsPVlRX4fttfUn69U0cM0n3zmAn41hhzxm3qPCsWiXZoiLdYReCzezfwvkzhb9lGKZJnRvlEEEMI1kmkIJ3R6AT6Fob6opl79pt4/Os+hZWzUpOplEzEJbIonQeEPy2c3emwpemHsBaV9WeEMytLgJMOXtjT17rSo9en2coSsRRa+qE5R/7Evv5pv0r9H1QPJr/AGbfv0YfSo9zi+0coSdR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nXl79ud6/IZ9CfU68vftzvX5DPoVfaZPpHFABFP/AOp15e/bnevyGfQn1OvL37c71+Qz6E9pk+jigAuWg7FP1vg68vQfKxnej+Kz6F31p2BMmKGRstwqLxcN3mySo3AfyVPtMieEq4mhztGtG8XctPoXucCZI5n5j1DYcLYSrp4idHVEkZjiaO0udwPo1VmGFtnDJbCO5JaMB27pWcpZ2dK8+cuWyKWio6GFtPR0sUETBo1kbA1oHmC2r0P/ALOEocZR7Alpt3Q3bNa6CvnB3hbaQkRN7nu5u9GgUvbHh6yYbtsNosNrpqCjgbuRwwRhjWjs0C/fuN+xC5Lrpirj+K8RpjdA5BZRFqsIiICIiAiIgIiIOO63sCzujqWUQEREBY0HYsogwmgHJZRRo04lvDTQLGg62hc0TW/I4ljSNC0EL881rtlR/CLfTS/fxNd8oX6kTUDqnYVww/2eHbYfPSR/QuTMM4di06Ow25mn2NLGP8l2aKOMfQ+ENHSQfvFNFHp9gwD5F9iB2BZRIjQ4huizoOzmsorAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiDHDkmnUFlFEb/oxqmqaDsTQdiaDVNU0HYmg7E0HBOHcmg7E0HYmg0HYsoikEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAR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oiAiIgLGgHBZRBjQdiaDnpyWUUan7GNBppp3LKIpBERAWNB2LKIMaDsWURAREQY3RrrogaBpoOSyiAsFoOuo5rKIMbo7FlEQEREBERBjQdiaDsWUQEREGNAU0HYsogIiICIiDG6OwJujlosogIiICIiAsaDs7llEGNB2LKIgIiIPyV9UaKhnqtxz+hifJutGpdoOQ71WtgzJ/MfNfP2nvWMsIXWiobvepLhXTTwFrGQ77nubqe5oaO4qzMgEaELG437ELLJijJPdExthkccbWxxtDWtAAA6tBp8i5oi1iNJEREBERAREQEREBERBx0aOOgTUdyzoOxNB2KvENAm6OzuWUUxv+giIpBYAA5BZRBjQdiyiICIiAsaBZRBggHmFlEQEREGA0A6gLKIgIiICIiAiIgxoOxZREBERR3BYLQTqQsoncERFIIiICIiAiIgLG6OxZRBjh3LGjVnQdiaDsUaDQdncsoikY0HZyTQa66LKICIiAiIgIiICIiDAAHILKIo0CxosoncY0WURSCIiDG6B1ck3R2BZRAREQEREBY3RpposogxoNNNFlEQFjdbrqAFlEGNB2dyyiICIiDG6NddE0HZzWUQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERB//2Q==")
                fname = f"{sel_ref}.xlsx"
                st.download_button(label="⬇️ Download Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.info("No past deliveries yet.")

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

    # ── Past Purchase Orders ─────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">📂 Past Purchase Orders</div>', unsafe_allow_html=True)
    log_df = load_log(SPREADSHEET_ID)
    if not log_df.empty:
        po_log = log_df[log_df["TXN_TYPE"] == "PO"]
        if po_log.empty:
            st.info("No past purchase orders yet.")
        else:
            past_po_refs = sorted(po_log["REF_NUMBER"].unique().tolist(), reverse=True)
            sel_po_ref = st.selectbox("Select PO Reference", past_po_refs, key="past_po_ref")
            if sel_po_ref:
                ref_rows = po_log[po_log["REF_NUMBER"] == sel_po_ref]
                first = ref_rows.iloc[0]
                dept_cols = ["RESTAURANT","BANQUET","CAFE","BAR","OTHERS"]
                def get_dept(r):
                    for dc in dept_cols:
                        if num(r.get(dc,0)) > 0:
                            return dc.title(), num(r.get(dc,0))
                    return "Others", 0
                total_val = 0
                st.markdown(f'<div class="info-box">📋 <strong>{sel_po_ref}</strong> &nbsp;|&nbsp; 📅 {first.get("DATE","")} &nbsp;|&nbsp; 👤 {first.get("STAFF","")} &nbsp;|&nbsp; {len(ref_rows)} item(s)</div>', unsafe_allow_html=True)

                table_rows = ""
                for row_idx, (_, r) in enumerate(ref_rows.iterrows(), 1):
                    dept_name, qty = get_dept(r)
                    unit_cost = num(items_df[items_df["ITEM"]==r["ITEM"]]["UNIT COST"].values[0]) if r["ITEM"] in items_df["ITEM"].values else 0
                    uom = items_df[items_df["ITEM"]==r["ITEM"]]["UNIT OF MEASURE"].values[0] if r["ITEM"] in items_df["ITEM"].values else ""
                    val = qty * unit_cost
                    total_val += val
                    st.markdown(f'<div class="po-item-row"><strong>{r["ITEM"]}</strong> &nbsp;|&nbsp; <span style="color:#8CAF7A;">{qty:,.2f} {uom}</span> &nbsp;|&nbsp; ₱{unit_cost:.4f}/unit &nbsp;|&nbsp; ₱{val:,.2f} &nbsp;|&nbsp; → {dept_name}</div>', unsafe_allow_html=True)
                    table_rows += f"<tr><td>{row_idx}</td><td>{r['ITEM']}</td><td>{uom}</td><td class='right'>{qty:,.2f}</td><td class='right'>&#8369;{unit_cost:.4f}</td><td class='right'>&#8369;{val:,.2f}</td><td>{dept_name}</td></tr>"

                st.markdown(f'<div class="info-box" style="text-align:right;">Total Value: <strong>₱{total_val:,.2f}</strong></div>', unsafe_allow_html=True)

                dept_display = sel_po_ref.split("-")[2].title() if len(sel_po_ref.split("-")) > 2 else "—"
                po_rows = []
                for ri, (_, r) in enumerate(ref_rows.iterrows(), 1):
                    dept_name, qty = get_dept(r)
                    unit_cost = num(items_df[items_df["ITEM"]==r["ITEM"]]["UNIT COST"].values[0]) if r["ITEM"] in items_df["ITEM"].values else 0
                    uom = items_df[items_df["ITEM"]==r["ITEM"]]["UNIT OF MEASURE"].values[0] if r["ITEM"] in items_df["ITEM"].values else ""
                    po_rows.append((ri, r["ITEM"], uom, float(qty), qty * unit_cost))
                buf = build_doc_xlsx("PURCHASE ORDER", sel_po_ref, first.get("DATE",""), first.get("STAFF",""), "Department", dept_display, po_rows, "/9j/4QA4RXhpZgAASUkqAAgAAAABAGmHBAABAAAAHAAAAAAAAAAAAAEAAJAHAAQAAAAwMjEwAAAAAAAA/+AAEEpGSUYAAQEAAAEAAQAA/+IB2ElDQ19QUk9GSUxFAAEBAAAByAAAAAAEMAAAbW50clJHQiBYWVogB+AAAQABAAAAAAAAYWNzcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAPbWAAEAAAAA0y0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJZGVzYwAAAPAAAAAkclhZWgAAARQAAAAUZ1hZWgAAASgAAAAUYlhZWgAAATwAAAAUd3RwdAAAAVAAAAAUclRSQwAAAWQAAAAoZ1RSQwAAAWQAAAAoYlRSQwAAAWQAAAAoY3BydAAAAYwAAAA8bWx1YwAAAAAAAAABAAAADGVuVVMAAAAIAAAAHABzAFIARwBCWFlaIAAAAAAAAG+iAAA49QAAA5BYWVogAAAAAAAAYpkAALeFAAAY2lhZWiAAAAAAAAAkoAAAD4QAALbPWFlaIAAAAAAAAPbWAAEAAAAA0y1wYXJhAAAAAAAEAAAAAmZmAADypwAADVkAABPQAAAKWwAAAAAAAAAAbWx1YwAAAAAAAAABAAAADGVuVVMAAAAgAAAAHABHAG8AbwBnAGwAZQAgAEkAbgBjAC4AIAAyADAAMQA2/9sAQwADAgICAgIDAgICAwMDAwQGBAQEBAQIBgYFBgkICgoJCAkJCgwPDAoLDgsJCQ0RDQ4PEBAREAoMEhMSEBMPEBAQ/9sAQwEDAwMEAwQIBAQIEAsJCxAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ/8AAEQgBkwcCAwEiAAIRAQMRAf/EAB0AAQACAgMBAQAAAAAAAAAAAAAICQEHAgUGAwT/xABnEAABAwIDBAQFCwwOBgcIAwEBAAIDBAUGBxEIEiExE0FRYQkUInGBFRgyN0J1kaGxs9EZI1JWYnJzgpKTlLIWFyQzNDU2OFNXdHa0wSVDVaKj4WSDwsPS8PEmJ0RFRlRjhChHpGX/xAAZAQEAAwEBAAAAAAAAAAAAAAAAAQIDBAX/xAAsEQEAAgIBAwQDAQEBAAMAAwAAAQIDERIEITETFDJRIjNBYUJSI3GBJDSR/9oADAMBAAIRAxEAPwC1NERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERARY1TTvUROxlFhE3AyiIm4BFj4UTlAyiIpBERAREQEREBERAREQEREBERAREQERYQZRflra+lt8DqqsqoaeFg8qSV4a0ek8FqDHm1vkrgPfhqMTtudWznTW5vTP8ASeAHwqlr1r5lG26UUJ8S+EZgjc6PCuAjJp7F9bPpr+K3T5VrW97fWddxc71MZabYw8hHTb5H4ztVhbq8df6jlCyJFVtNtn7QEr944yaw/YtpmNHyKRmydtZ4pzJxOMvcfMhnrZYny0lbEwML91pcWOA4chzU06ql7cSLbS9WD59EKr+2odprODDWbl3wnhq+zWe22oxRQthaAZdYw4vLjz4kjT7la5ckYq8pTM6WAhwPJZVYFh22c+bLM18uJILjEOcdXTNdr6dNVuLCXhFXno4saYIb91NQTHQ/iu1WNerpZWLxKbqLTGCNrbJTHDWMpsVR22qd/wDD3AdC70HkfhW2bdeLddoPGrXcKashPJ8Eoe0+kLeuStvErbfuRY4r81fWQ2+inraqdsMUEbnve9wAaAOZJV0vt0g9j0g15ceeq+iqSxfnRjy75m3DFVBjC5DW4yOpSypcGMj3/JaBrpporWsPVlRX4fttfUn69U0cM0n3zmAn41hhzxm3qPCsWiXZoiLdYReCzezfwvkzhb9lGKZJnRvlEEEMI1kmkIJ3R6AT6Fob6opl79pt4/Os+hZWzUpOplEzEJbIonQeEPy2c3emwpemHsBaV9WeEMytLgJMOXtjT17rSo9en2coSsRRa+qE5R/7Evv5pv0r9H1QPJr/AGbfv0YfSo9zi+0coSdR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nXl79ud6/IZ9CfU68vftzvX5DPoVfaZPpHFABFP/AOp15e/bnevyGfQn1OvL37c71+Qz6E9pk+jigAuWg7FP1vg68vQfKxnej+Kz6F31p2BMmKGRstwqLxcN3mySo3AfyVPtMieEq4mhztGtG8XctPoXucCZI5n5j1DYcLYSrp4idHVEkZjiaO0udwPo1VmGFtnDJbCO5JaMB27pWcpZ2dK8+cuWyKWio6GFtPR0sUETBo1kbA1oHmC2r0P/ALOEocZR7Alpt3Q3bNa6CvnB3hbaQkRN7nu5u9GgUvbHh6yYbtsNosNrpqCjgbuRwwRhjWjs0C/fuN+xC5Lrpirj+K8RpjdA5BZRFqsIiICIiAiIgIiIOO63sCzujqWUQEREBY0HYsogwmgHJZRRo04lvDTQLGg62hc0TW/I4ljSNC0EL881rtlR/CLfTS/fxNd8oX6kTUDqnYVww/2eHbYfPSR/QuTMM4di06Ow25mn2NLGP8l2aKOMfQ+ENHSQfvFNFHp9gwD5F9iB2BZRIjQ4huizoOzmsorAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiDHDkmnUFlFEb/oxqmqaDsTQdiaDVNU0HYmg7E0HBOHcmg7E0HYmg0HYsoikEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAR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oiAiIgLGgHBZRBjQdiaDnpyWUUan7GNBppp3LKIpBERAWNB2LKIMaDsWURAREQY3RrrogaBpoOSyiAsFoOuo5rKIMbo7FlEQEREBERBjQdiaDsWUQEREGNAU0HYsogIiICIiDG6OwJujlosogIiICIiAsaDs7llEGNB2LKIgIiIPyV9UaKhnqtxz+hifJutGpdoOQ71WtgzJ/MfNfP2nvWMsIXWiobvepLhXTTwFrGQ77nubqe5oaO4qzMgEaELG437ELLJijJPdExthkccbWxxtDWtAAA6tBp8i5oi1iNJEREBERAREQEREBERBx0aOOgTUdyzoOxNB2KvENAm6OzuWUUxv+giIpBYAA5BZRBjQdiyiICIiAsaBZRBggHmFlEQEREGA0A6gLKIgIiICIiAiIgxoOxZREBERR3BYLQTqQsoncERFIIiICIiAiIgLG6OxZRBjh3LGjVnQdiaDsUaDQdncsoikY0HZyTQa66LKICIiAiIgIiICIiDAAHILKIo0CxosoncY0WURSCIiDG6B1ck3R2BZRAREQEREBY3RpposogxoNNNFlEQFjdbrqAFlEGNB2dyyiICIiDG6NddE0HZzWUQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERB//2Q==")
                fname = f"{sel_po_ref}.xlsx"
                st.download_button(label="⬇️ Download Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.info("No past purchase orders yet.")

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
            total_row = 650
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

        total_row = 650
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