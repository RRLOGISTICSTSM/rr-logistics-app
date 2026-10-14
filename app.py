import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io
import tempfile
import sqlite3

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Safe import for openpyxl
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

st.set_page_config(page_title="R.R. Logistics TSM", page_icon="🚚", layout="wide")

# ==========================================
# 0. DATABASE & STORAGE SETUP
# ==========================================
UPLOAD_FOLDER = "saved_documents"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def init_db():
    conn = sqlite3.connect("rr_logistics.db")
    c = conn.cursor()
    
    # ॲप सेटिंग्ज आणि पासवर्ड कायमस्वरूपी सेव्ह ठेवण्यासाठी
    c.execute('''CREATE TABLE IF NOT EXISTS app_settings (
                    key_name TEXT PRIMARY KEY,
                    key_value TEXT
                )''')
    c.execute("INSERT OR IGNORE INTO app_settings VALUES ('admin_password', 'rr123')")
    c.execute("INSERT OR IGNORE INTO app_settings VALUES ('security_pin', '9822')")

    # वाहनांचे टेबल
    c.execute('''CREATE TABLE IF NOT EXISTS vehicles (
                    veh_no TEXT PRIMARY KEY,
                    owner_name TEXT,
                    veh_type TEXT,
                    capacity TEXT,
                    default_dl TEXT,
                    rc_doc TEXT,
                    ins_doc TEXT,
                    ins_exp TEXT,
                    fit_doc TEXT,
                    fit_exp TEXT,
                    puc_doc TEXT,
                    puc_exp TEXT,
                    per_doc TEXT,
                    per_exp TEXT,
                    created_at TEXT
                )''')

    # ड्रायव्हर्स टेबल
    c.execute('''CREATE TABLE IF NOT EXISTS drivers (
                    dl_no TEXT PRIMARY KEY,
                    driver_name TEXT,
                    mobile TEXT,
                    aadhaar_no TEXT,
                    upi_id TEXT,
                    bank_name TEXT,
                    acc_no TEXT,
                    ifsc TEXT,
                    aadhaar_doc TEXT,
                    lic_doc TEXT,
                    created_at TEXT
                )''')

    # ट्रिप शीट्स टेबल
    c.execute('''CREATE TABLE IF NOT EXISTS trip_sheets (
                    lr_no TEXT PRIMARY KEY,
                    loading_date TEXT,
                    date_val TEXT,
                    from_loc TEXT,
                    to_loc TEXT,
                    truck_no TEXT,
                    driver_dl TEXT,
                    driver_name TEXT,
                    driver_mob TEXT,
                    rate_ton REAL,
                    weight_ton REAL,
                    freight REAL,
                    bill_no TEXT,
                    otp TEXT,
                    party_name TEXT,
                    material_grade TEXT,
                    start_km INTEGER,
                    end_km INTEGER,
                    running_km INTEGER,
                    trip_adv REAL,
                    unloading REAL,
                    food REAL,
                    tyre_bill REAL,
                    fastag1 REAL,
                    fastag2 REAL,
                    diesel_ltr REAL,
                    diesel_amt REAL,
                    commission REAL,
                    maintenance REAL,
                    other_exp REAL,
                    total_exp REAL,
                    trip_balance REAL,
                    created_at TEXT
                )''')

    # सुरक्षित Schema Migration - जुना डेटा डिलीट न होता नवीन कॉलम्स जोडणे
    c.execute("PRAGMA table_info(trip_sheets)")
    existing_cols = [col[1] for col in c.fetchall()]
    if "di_no" not in existing_cols:
        try:
            c.execute("ALTER TABLE trip_sheets ADD COLUMN di_no TEXT DEFAULT ''")
        except:
            pass
    if "bolt_seal_no" not in existing_cols:
        try:
            c.execute("ALTER TABLE trip_sheets ADD COLUMN bolt_seal_no TEXT DEFAULT ''")
        except:
            pass
    if "material_grade" not in existing_cols:
        try:
            c.execute("ALTER TABLE trip_sheets ADD COLUMN material_grade TEXT DEFAULT 'PPC'")
        except:
            pass

    conn.commit()
    conn.close()

init_db()

def get_db_connection():
    return sqlite3.connect("rr_logistics.db")

def get_setting(key, default_val=""):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT key_value FROM app_settings WHERE key_name=?", (key,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else default_val

def set_setting(key, val):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO app_settings VALUES (?, ?)", (key, str(val)))
    conn.commit()
    conn.close()

def fetch_vehicle_list():
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT veh_no FROM vehicles ORDER BY veh_no", conn)
    conn.close()
    return df['veh_no'].tolist()

def fetch_driver_list():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT dl_no, driver_name FROM drivers ORDER BY driver_name")
    rows = c.fetchall()
    conn.close()
    return rows

def get_vehicle_data(veh_no):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM vehicles WHERE veh_no=?", (veh_no,))
    row = c.fetchone()
    conn.close()
    return row

def get_driver_data(dl_no):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM drivers WHERE dl_no=?", (dl_no,))
    row = c.fetchone()
    conn.close()
    return row

def save_uploaded_file(uploaded_file, prefix, existing_path=""):
    if uploaded_file is None:
        return existing_path
    file_ext = os.path.splitext(uploaded_file.name)[1]
    safe_name = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_ext}"
    file_path = os.path.join(UPLOAD_FOLDER, safe_name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

# ==========================================
# 1. LR PDF GENERATION FUNCTION
# ==========================================
def generate_official_lr(lr_no, di_no, bolt_no, truck_no, date_val, from_loc, to_loc, consignor, consignee, c_gst, c_eway, d_name, d_dl, d_mobile, weight, freight, order_by, mat_grade):
    temp_dir = tempfile.gettempdir()
    safe_lr = str(lr_no).replace('/', '_')
    pdf_path = os.path.join(temp_dir, f"LR_{safe_lr}.pdf")
    
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=2)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=2)
    addr_style = ParagraphStyle('AddrStyle', parent=styles['Normal'], fontSize=7.5, textColor=colors.HexColor('#334155'), alignment=1, spaceAfter=8)
    
    for l_file in ["logo.jpeg", "logo.jpg", "logo.png"]:
        if os.path.exists(l_file):
            img = RLImage(l_file, width=80, height=55)
            img.hAlign = 'CENTER'
            story.append(img)
            story.append(Spacer(1, 4))
            break
    
    story.append(Paragraph("SUBJECT TO DHULE JURISDICTION", ParagraphStyle('SubTop', parent=styles['Normal'], fontSize=6.5, textColor=colors.HexColor('#475569'), alignment=1)))
    story.append(Paragraph("R.R. LOGISTICS TRANSPORT PVT. LTD.", title_style))
    story.append(Paragraph("TRANSPORT CONTRACTOR & COMMISSION AGENT", sub_style))
    story.append(Paragraph("Gat No. 298-2A, R.R. Petroleum, Mumbai-Agra Highway, Sarwad, Dhule.<br/><b>Mob: 9822414040 / 9049521177</b>", addr_style))
    
    data = [
        [
            Paragraph(f"<b>L.R. No. :</b> {lr_no}", styles['Normal']),
            Paragraph(f"<b>DI No. :</b> {di_no}<br/><b>Bolt / Seal No. :</b> {bolt_no}", styles['Normal']),
            Paragraph(f"<b>Date :</b> {date_val}<br/><b>Truck No. :</b> {truck_no}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>From :</b> {from_loc}", styles['Normal']),
            "",
            Paragraph(f"<b>To :</b> {to_loc}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>Consignor:</b><br/>{consignor}<br/><br/><b>GSTIN No. :</b> {c_gst}<br/><b>E-way Bill No. :</b> {c_eway}", styles['Normal']),
            "",
            Paragraph(f"<b>Consignee:</b><br/>{consignee}<br/><br/><b>Material / Grade:</b> {mat_grade}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>Driver Name :</b> {d_name}", styles['Normal']),
            Paragraph(f"<b>D.L. No. :</b> {d_dl}", styles['Normal']),
            Paragraph(f"<b>Weight (MT) :</b> {weight}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>Driver Mobile :</b> {d_mobile}", styles['Normal']),
            "",
            Paragraph(f"<b>Freight (भाडे) :</b> {freight}<br/><b>Order By :</b> {order_by}", styles['Normal'])
        ],
        [
            Paragraph("We are only Broker and commission agent. Please check all documents of truck carefully. Subject to DHULE Jurisdiction.", styles['Normal']),
            "",
            Paragraph("For R.R. LOGISTICS TRANSPORT PVT. LTD.<br/><br/><br/><b>Authorized Signatory</b>", styles['Normal'])
        ]
    ]
    
    t = Table(data, colWidths=[200, 185, 185])
    t.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#1e3a8a')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#1e3a8a')),
        ('SPAN', (0, 1), (1, 1)),
        ('SPAN', (0, 2), (1, 2)),
        ('SPAN', (0, 4), (1, 4)),
        ('SPAN', (0, 5), (1, 5)),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    
    story.append(t)
    story.append(Spacer(1, 4))
    story.append(Paragraph("THIS IS COMPUTER GENERATED BILTY", ParagraphStyle('Foot', parent=styles['Normal'], fontSize=6.5, textColor=colors.HexColor('#475569'), alignment=1)))
    
    doc.build(story)
    return pdf_path

# ==========================================
# 2. TRIP SHEET EXCEL (WITH FORMULAS)
# ==========================================
def generate_trip_sheet_excel(d):
    if not HAS_OPENPYXL:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trip Sheet"
    
    navy_fill = PatternFill(start_color="103A70", end_color="103A70", fill_type="solid")
    blue_header_fill = PatternFill(start_color="95B3D7", end_color="95B3D7", fill_type="solid")
    dark_blue_block = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = "R. R. LOGISTICS, DHULE - TRIP SHEET"
    title_cell.font = font_title
    title_cell.fill = navy_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    headers_meta = [
        ("A2", "LR.NO", "B2", d['lr_no'], "C2", "DI NO", "D2", d.get('di_no', ''), "E2", "Date", "F2", d['date_val']),
        ("A3", "From", "B3", d['from_loc'], "C3", "To", "D3", d['to_loc'], "E3", "Truck No", "F3", d['truck_no']),
        ("A4", "Rate (per Ton)", "B4", d['rate_ton'], "C4", "Freight", "D4", "=B4*F4", "E4", "Weight (Tons)", "F4", d['weight_ton']),
        ("A5", "Bolt/Seal No", "B5", d.get('bolt_seal_no', ''), "C5", "Grade", "D5", d.get('material_grade', 'PPC'), "E5", "Party Name", "F5", d['party_name'])
    ]
    
    for row in headers_meta:
        ws[row[0]] = row[1]; ws[row[0]].font = font_bold
        ws[row[2]] = row[3]; ws[row[2]].font = font_regular
        ws[row[4]] = row[5]; ws[row[4]].font = font_bold
        ws[row[6]] = row[7]; ws[row[6]].font = font_regular
        ws[row[8]] = row[9]; ws[row[8]].font = font_bold
        ws[row[10]] = row[11]; ws[row[10]].font = font_regular
        
    ws.merge_cells("G2:H5")
    for r in range(2, 6):
        for c in range(7, 9):
            ws.cell(row=r, column=c).fill = dark_blue_block

    exp_headers = [("A7", "EXP"), ("B7", "Ltr"), ("C7", "Rate"), ("D7", "Amount"), ("E7", "Date"), ("F7", "Pump"), ("G7", "Summary"), ("H7", "Values")]
    for cell, val in exp_headers:
        ws[cell] = val
        ws[cell].fill = blue_header_fill
        ws[cell].font = font_bold
        ws[cell].alignment = Alignment(horizontal="center")

    exp_data = [
        ("Trip Adv", "", "", d['trip_adv'], d['adv_date'], ""),
        ("Unloading", "", "", d['unloading'], "", ""),
        ("Food", "", "", d['food'], "", ""),
        ("Tyre Bill", "", "", d['tyre_bill'], "", ""),
        ("FasTag-1", "", "", d['fastag1'], "", ""),
        ("FasTag-2", "", "", d['fastag2'], "", ""),
        ("Diesel-1", d['d1_ltr'], d['d1_rate'], "=B14*C14", d['d1_date'], d['d1_pump']),
        ("Diesel-2", d['d2_ltr'], d['d2_rate'], "=B15*C15", d['d2_date'], d['d2_pump']),
        ("Commission", "", "", d['commission'], "", ""),
        ("Maintenance", "", "", d['maintenance'], "", ""),
        ("Other", "", "", d['other_exp'], "", "")
    ]

    for idx, r_data in enumerate(exp_data, start=8):
        ws[f"A{idx}"] = r_data[0]
        ws[f"B{idx}"] = r_data[1]
        ws[f"C{idx}"] = r_data[2]
        ws[f"D{idx}"] = r_data[3]
        ws[f"E{idx}"] = r_data[4]
        ws[f"F{idx}"] = r_data[5]

    ws["A19"] = "Total Exp"
    ws["A19"].font = font_bold
    ws["D19"] = "=SUM(D8:D18)"
    ws["D19"].font = font_bold

    summary_data = [
        ("Start Km", d['start_km']),
        ("End Km", d['end_km']),
        ("Running Km", "=H9-H8"),
        ("Total Diesel (Ltr)", "=B14+B15"),
        ("Average (Km/L)", "=IF(H11>0, H10/H11, 0)"),
        ("Trip Balance", "=D4-D19"),
    ]
    for idx, (lbl, val) in enumerate(summary_data, start=8):
        ws[f"G{idx}"] = lbl
        ws[f"H{idx}"] = val
        ws[f"G{idx}"].font = font_bold

    ws["G19"] = "Trip Balance"
    ws["G19"].font = font_bold
    ws["H19"] = "=D4-D19"
    ws["H19"].font = font_bold

    for r in list(range(8, 14)) + list(range(16, 19)):
        ws[f"B{r}"].fill = dark_blue_block
        ws[f"C{r}"].fill = dark_blue_block

    for row in ws.iter_rows(min_row=1, max_row=19, min_col=1, max_col=8):
        for cell in row:
            cell.border = thin_border
            if cell.row in [2,3,4,5,7]:
                if not cell.alignment.horizontal:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            if str(cell.coordinate)[0] in ['B', 'C', 'D', 'H'] and cell.row >= 8:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ==========================================
# 3. TRIP SHEET PDF GENERATION
# ==========================================
def generate_trip_sheet_pdf(d):
    temp_dir = tempfile.gettempdir()
    safe_lr = str(d['lr_no']).replace('/', '_')
    pdf_path = os.path.join(temp_dir, f"TripSheet_{safe_lr}.pdf")
    doc = SimpleDocTemplate(pdf_path, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    
    freight_val = round(d['rate_ton'] * d['weight_ton'], 2)
    d1_amt = round(d['d1_ltr'] * d['d1_rate'], 2)
    d2_amt = round(d['d2_ltr'] * d['d2_rate'], 2)
    tot_exp = round(d['trip_adv'] + d['unloading'] + d['food'] + d['tyre_bill'] + d['fastag1'] + d['fastag2'] + d1_amt + d2_amt + d['commission'] + d['maintenance'] + d['other_exp'], 2)
    running_km = d['end_km'] - d['start_km']
    total_diesel = round(d['d1_ltr'] + d['d2_ltr'], 2)
    avg_kml = round(running_km / total_diesel, 2) if total_diesel > 0 else 0
    trip_bal = round(freight_val - tot_exp, 2)

    table_data =
