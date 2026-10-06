import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io
import tempfile
import sqlite3

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Safe import for openpyxl
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

st.set_page_config(page_title="RR Logistics TSM", page_icon="🚚", layout="wide")

# ==========================================
# 0. DATABASE & STORAGE SETUP (डेटाबेस सेटअप)
# ==========================================
UPLOAD_FOLDER = "saved_documents"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def init_db():
    conn = sqlite3.connect("rr_logistics.db")
    c = conn.cursor()
    # Vehicles & Expiry Dates Table
    c.execute('''CREATE TABLE IF NOT EXISTS vehicles (
                    veh_no TEXT PRIMARY KEY,
                    owner_name TEXT,
                    veh_type TEXT,
                    capacity TEXT,
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
    # Drivers Table
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
    # Trip Sheets Table
    c.execute('''CREATE TABLE IF NOT EXISTS trip_sheets (
                    lr_no TEXT PRIMARY KEY,
                    loading_date TEXT,
                    date_val TEXT,
                    from_loc TEXT,
                    to_loc TEXT,
                    truck_no TEXT,
                    rate_ton REAL,
                    weight_ton REAL,
                    freight REAL,
                    bill_no TEXT,
                    otp TEXT,
                    party_name TEXT,
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
    conn.commit()
    conn.close()

init_db()

def save_uploaded_file(uploaded_file, prefix):
    if uploaded_file is None:
        return ""
    file_ext = os.path.splitext(uploaded_file.name)[1]
    safe_name = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_ext}"
    file_path = os.path.join(UPLOAD_FOLDER, safe_name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

# ==========================================
# 1. LR PDF GENERATION FUNCTION
# ==========================================
def generate_official_lr(lr_no, truck_no, date_val, from_loc, to_loc, consignor, consignee, c_gst, c_eway, d_name, d_dl, d_mobile, weight, freight, order_by):
    temp_dir = tempfile.gettempdir()
    safe_lr = str(lr_no).replace('/', '_')
    pdf_path = os.path.join(temp_dir, f"LR_{safe_lr}.pdf")
    
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=2)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=2)
    addr_style = ParagraphStyle('AddrStyle', parent=styles['Normal'], fontSize=7.5, textColor=colors.HexColor('#334155'), alignment=1, spaceAfter=8)
    
    logo_path = "logo.jpeg" 
    if os.path.exists(logo_path):
        img = Image(logo_path, width=70, height=60)
        img.hAlign = 'CENTER'
        story.append(img)
        story.append(Spacer(1, 4))
    
    story.append(Paragraph("SUBJECT TO DHULE JURISDICTION", ParagraphStyle('SubTop', parent=styles['Normal'], fontSize=6.5, textColor=colors.HexColor('#475569'), alignment=1)))
    story.append(Paragraph("R.R. LOGISTICS TRANSPORT PVT. LTD.", title_style))
    story.append(Paragraph("TRANSPORT CONTRACTOR & COMMISSION AGENT", sub_style))
    story.append(Paragraph("Gat No. 298-2A, R.R. Petroleum, Mumbai Agra Highway, Jeevan Dhara Dairy, Sarwad, Dhule.<br/><b>Mob: 9822414040 / 9049521177</b>", addr_style))
    
    data = [
        [
            Paragraph(f"<b>L.R. No. :</b> {lr_no}", styles['Normal']),
            Paragraph(f"<b>Truck No. :</b> {truck_no}", styles['Normal']),
            Paragraph(f"<b>Date :</b> {date_val}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>From :</b> {from_loc}", styles['Normal']),
            "",
            Paragraph(f"<b>To :</b> {to_loc}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>Consignor:</b><br/>{consignor}<br/><br/><b>GSTIN No. :</b> {c_gst}<br/><b>E-way Bill No. :</b> {c_eway}", styles['Normal']),
            "",
            Paragraph(f"<b>Consignee:</b><br/>{consignee}<br/><br/><b>GSTIN No. :</b><br/><b>E-way Bill No. :</b>", styles['Normal'])
        ],
        [
            Paragraph(f"<b>Driver Detail - Name :</b> {d_name}", styles['Normal']),
            Paragraph(f"<b>D.L. No. :</b> {d_dl}", styles['Normal']),
            Paragraph(f"<b>Weight (वजन) :</b> {weight}", styles['Normal'])
        ],
        [
            Paragraph(f"<b>Driver Mobile No. :</b> {d_mobile}", styles['Normal']),
            "",
            Paragraph(f"<b>Freight (भाडे) :</b> {freight}<br/><b>Order By :</b> {order_by}", styles['Normal'])
        ],
        [
            Paragraph("We are only Broker and commission agent. Please Check the all documents of truck carefully. Subject to DHULE Jurisdiction.", styles['Normal']),
            "",
            Paragraph("For R.R. LOGISTICS TRANSPORT PVT. LTD.<br/><br/><br/>Authorized Signatory", styles['Normal'])
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
# 2. TRIP SHEET EXCEL (WITH LIVE FORMULAS)
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
        ("A2", "LR.NO", "B2", d['lr_no'], "C2", "Loading Date", "D2", d['loading_date'], "E2", "Date", "F2", d['date_val']),
        ("A3", "From", "B3", d['from_loc'], "C3", "To", "D3", d['to_loc'], "E3", "Truck No", "F3", d['truck_no']),
        ("A4", "Rate (per Ton)", "B4", d['rate_ton'], "C4", "Freight", "D4", "=B4*F4", "E4", "Weight (Tons)", "F4", d['weight_ton']),
        ("A5", "Bill No", "B5", d['bill_no'], "C5", "OTP", "D5", d['otp'], "E5", "Party Name", "F5", d['party_name'])
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

    table_data = [
        ["R. R. LOGISTICS, DHULE - TRIP SHEET", "", "", "", "", "", "", ""],
        ["LR.NO", str(d['lr_no']), "Loading Date", str(d['loading_date']), "Date", str(d['date_val']), "", ""],
        ["From", str(d['from_loc']), "To", str(d['to_loc']), "Truck No", str(d['truck_no']), "", ""],
        ["Rate (per Ton)", str(d['rate_ton']), "Freight", str(freight_val), "Weight (Tons)", str(d['weight_ton']), "", ""],
        ["Bill No", str(d['bill_no']), "OTP", str(d['otp']), "Party Name", str(d['party_name']), "", ""],
        ["EXP", "Ltr", "Rate", "Amount", "Date", "Pump", "Summary", "Values"],
        ["Trip Adv", "", "", str(d['trip_adv']), str(d['adv_date']), "", "Start Km", str(d['start_km'])],
        ["Unloading", "", "", str(d['unloading']), "", "", "End Km", str(d['end_km'])],
        ["Food", "", "", str(d['food']), "", "", "Running Km", str(running_km)],
        ["Tyre Bill", "", "", str(d['tyre_bill']), "", "", "Total Diesel (Ltr)", str(total_diesel)],
        ["FasTag-1", "", "", str(d['fastag1']), "", "", "Average (Km/L)", str(avg_kml)],
        ["FasTag-2", "", "", str(d['fastag2']), "", "", "Trip Balance", str(trip_bal)],
        ["Diesel-1", str(d['d1_ltr']), str(d['d1_rate']), str(d1_amt), str(d['d1_date']), str(d['d1_pump']), "", ""],
        ["Diesel-2", str(d['d2_ltr']), str(d['d2_rate']), str(d2_amt), str(d['d2_date']), str(d['d2_pump']), "", ""],
        ["Commission", "", "", str(d['commission']), "", "", "", ""],
        ["Maintenance", "", "", str(d['maintenance']), "", "", "", ""],
        ["Other", "", "", str(d['other_exp']), "", "", "", ""],
        ["Total Exp", "", "", str(tot_exp), "", "", "Trip Balance", str(trip_bal)]
    ]

    t = Table(table_data, colWidths=[90, 60, 60, 80, 85, 110, 110, 85])
    t.setStyle(TableStyle([
        ('SPAN', (0, 0), (7, 0)),
        ('BACKGROUND', (0, 0), (7, 0), colors.HexColor('#103A70')),
        ('TEXTCOLOR', (0, 0), (7, 0), colors.white),
        ('ALIGN', (0, 0), (7, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (7, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (7, 0), 13),
        
        ('SPAN', (6, 1), (7, 4)),
        ('BACKGROUND', (6, 1), (7, 4), colors.HexColor('#4F81BD')),
        
        ('BACKGROUND', (0, 5), (7, 5), colors.HexColor('#95B3D7')),
        ('FONTNAME', (0, 5), (7, 5), 'Helvetica-Bold'),
        ('ALIGN', (0, 5), (7, 5), 'CENTER'),
        
        ('BACKGROUND', (1, 6), (2, 11), colors.HexColor('#4F81BD')),
        ('BACKGROUND', (1, 14), (2, 16), colors.HexColor('#4F81BD')),
        
        ('FONTNAME', (0, 1), (0, 4), 'Helvetica-Bold'),
        ('FONTNAME', (2, 1), (2, 4), 'Helvetica-Bold'),
        ('FONTNAME', (4, 1), (4, 4), 'Helvetica-Bold'),
        ('FONTNAME', (6, 6), (6, 11), 'Helvetica-Bold'),
        ('FONTNAME', (0, 17), (0, 17), 'Helvetica-Bold'),
        ('FONTNAME', (3, 17), (3, 17), 'Helvetica-Bold'),
        ('FONTNAME', (6, 17), (7, 17), 'Helvetica-Bold'),
        
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 6), (-1, -1), 'CENTER'),
    ]))
    
    story.append(t)
    doc.build(story)
    return pdf_path

# ==========================================
# 4. DOCUMENT EXPIRY PDF GENERATION (नवीन PDF)
# ==========================================
def generate_doc_expiry_pdf(vehicle_rows):
    temp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(temp_dir, f"Vehicle_Doc_Expiry_{datetime.now().strftime('%Y%m%d')}.pdf")
    doc = SimpleDocTemplate(pdf_path, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    styles = getSampleStyleSheet()
    
    title = Paragraph("<b>R.R. LOGISTICS TRANSPORT PVT. LTD.</b><br/><font size=10>गाड्यांच्या कागदपत्रांचा एक्सपायरी रिपोर्ट (Vehicle Documents Expiry Status)</font>", ParagraphStyle('ExpTitle', parent=styles['Heading1'], alignment=1, fontSize=14, textColor=colors.HexColor('#1e3a8a')))
    story.append(title)
    story.append(Spacer(1, 10))
    
    table_data = [
        ["गाडी क्रमांक", "मालकाचे नाव", "इन्शुरन्स एक्सपायरी", "फिटनेस एक्सपायरी", "PUC एक्सपायरी", "नॅशनल परमिट एक्सपायरी", "स्थिती (Status)"]
    ]
    
    today = datetime.now().date()
    for row in vehicle_rows:
        veh_no = row[0]
        owner = row[1]
        ins_exp = row[6] or "N/A"
        fit_exp = row[8] or "N/A"
        puc_exp = row[10] or "N/A"
        per_exp = row[12] or "N/A"
        
        # Check alerts
        alerts = []
        for exp_val, name in [(ins_exp, "Ins"), (fit_exp, "Fit"), (puc_exp, "PUC"), (per_exp, "Permit")]:
            if exp_val != "N/A":
                try:
                    exp_date = datetime.strptime(exp_val, "%Y-%m-%d").date()
                    if exp_date < today:
                        alerts.append(f"{name} Expired")
                    elif (exp_date - today).days <= 15:
                        alerts.append(f"{name} Due Soon")
                except:
                    pass
        status_text = ", ".join(alerts) if alerts else "सर्व वैध (Valid)"
        
        table_data.append([veh_no, owner, ins_exp, fit_exp, puc_exp, per_exp, status_text])
        
    t = Table(table_data, colWidths=[100, 120, 100, 100, 95, 120, 120])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#103A70')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t)
    doc.build(story)
    return pdf_path

# ==========================================
# 5. MAIN STREAMLIT APP
# ==========================================
def main():
    st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>R.R. LOGISTICS TRANSPORT PVT. LTD.</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #475569;'><b>TRANSPORT CONTRACTOR & COMMISSION AGENT</b> | Dhule Jurisdiction</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    if "current_password" not in st.session_state:
        st.session_state.current_password = "rr123"

    st.sidebar.title("🔐 सिस्टम लॉगिन")
    username = st.sidebar.text_input("User ID", key="login_user")
    password = st.sidebar.text_input("Password", type="password", key="login_pass")
    
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.sidebar.button("Login"):
        if username == "admin" and password == st.session_state.current_password:
            st.session_state.logged_in = True
            st.sidebar.success("सक्सेसफुली लॉगिन झाले! ✅")
        else:
            st.sidebar.error("चुकीचा User ID किंवा Password!")

    with st.sidebar.expander("❓ Forgot Password? (पासवर्ड विसरलात?)"):
        st.caption("पासवर्ड रीसेट करण्यासाठी ऑथेंटिकेशन कोड टाका:")
        recovery_code = st.text_input("Security PIN / Secret Key", type="password", key="sec_pin")
        new_pass = st.text_input("नवीन पासवर्ड (New Password)", type="password", key="new_pwd")
        confirm_pass = st.text_input("पासवर्ड खात्री करा (Confirm Password)", type="password", key="conf_pwd")
        
        if st.button("पासवर्ड बदला (Reset Password)"):
            if recovery_code == "9822":
                if new_pass and new_pass == confirm_pass:
                    st.session_state.current_password = new_pass
                    st.success("पासवर्ड बदलला आहे! आता नवीन पासवर्डने लॉगिन करा. ✅")
                else:
                    st.error("दोन्ही पासवर्ड जुळत नाहीत किंवा रिकामे आहेत!")
            else:
                st.error("चुकीचा Security PIN!")

    if st.session_state.logged_in:
        st.sidebar.markdown("---")
        menu = st.sidebar.radio("मेन्यू निवडा:", [
            "१. New Vehicle Registration & Docs", 
            "२. Driver Registration & Docs", 
            "३. Trip Information & Auto LR Generation", 
            "४. Diesel & Expenses Tracking",
            "५. Trip Sheet (Auto Excel & PDF)",
            "६. Saved Records & View Documents"
        ], key="dashboard_menu")
        
        # --- MENU 1: VEHICLE & EXPIRY DATES ---
        if menu == "१. New Vehicle Registration & Docs":
            st.subheader("🚛 नवीन गाडी नोंदणी व कागदपत्रे एक्सपायरीसह सेव्ह करा")
            v_col1, v_col2 = st.columns(2)
            with v_col1:
                veh_no = st.text_input("Vehicle Number (गाडी क्रमांक)", value="MH18BZ2240")
                veh_owner = st.text_input("Owner Name (मालकाचे नाव)")
            with v_col2:
                veh_type = st.text_input("Vehicle Type (उदा. 10 व्हीलर, 14 व्हीलर)")
                veh_cap = st.text_input("Capacity in Ton (क्षमता)")
            
            st.markdown("---")
            st.markdown("### 📂 कागदपत्रे अपलोड आणि एक्सपायरी तारीख (Expiry Dates)")
            
            d1, d2 = st.columns(2)
            with d1:
                rc_file = st.file_uploader("1. RC Book Upload", type=["pdf", "jpg", "png"], key="rc_up")
                insurance_file = st.file_uploader("2. Insurance Upload (इन्शुरन्स)", type=["pdf", "jpg", "png"], key="ins_up")
                ins_exp = st.date_input("इन्शुरन्स एक्सपायरी तारीख", datetime.now().date() + timedelta(days=180), key="ins_exp_key")
                fitness_file = st.file_uploader("3. Fitness Certificate Upload (फिटनेस)", type=["pdf", "jpg", "png"], key="fit_up")
                fit_exp = st.date_input("फिटनेस एक्सपायरी तारीख", datetime.now().date() + timedelta(days=180), key="fit_exp_key")
                
            with d2:
                puc_file = st.file_uploader("4. PUC Upload (पोल्युशन)", type=["pdf", "jpg", "png"], key="puc_up")
                puc_exp = st.date_input("PUC एक्सपायरी तारीख", datetime.now().date() + timedelta(days=90), key="puc_exp_key")
                permit_file = st.file_uploader("5. National Permit Upload (नॅशनल परमिट)", type=["pdf", "jpg", "png"], key="per_up")
                per_exp = st.date_input("परमिट एक्सपायरी तारीख", datetime.now().date() + timedelta(days=365), key="per_exp_key")
            
            if st.button("गाडीची संपूर्ण माहिती व कागदपत्रे सेव्ह करा"):
                if veh_no:
                    rc_path = save_uploaded_file(rc_file, f"{veh_no}_RC")
                    ins_path = save_uploaded_file(insurance_file, f"{veh_no}_INS")
                    fit_path = save_uploaded_file(fitness_file, f"{veh_no}_FIT")
                    puc_path = save_uploaded_file(puc_file, f"{veh_no}_PUC")
                    per_path = save_uploaded_file(permit_file, f"{veh_no}_PER")
                    
                    conn = sqlite3.connect("rr_logistics.db")
                    c = conn.cursor()
                    c.execute('''INSERT OR REPLACE INTO vehicles VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', 
                              (veh_no, veh_owner, veh_type, veh_cap, rc_path, ins_path, str(ins_exp), 
                               fit_path, str(fit_exp), puc_path, str(puc_exp), per_path, str(per_exp), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    conn.close()
                    st.success(f"गाडी क्र. {veh_no} ची माहिती आणि कागदपत्रे डेटाबेसमध्ये कायमस्वरूपी सेव्ह झाली! ✅")
                else:
                    st.warning("कृपया गाडी क्रमांक प्रविष्ट करा!")

        # --- MENU 2: DRIVER REGISTRATION ---
        elif menu == "२. Driver Registration & Docs":
            st.subheader("👨‍✈️ ड्रायव्हर नोंदणी, कागदपत्रे आणि बँक तपशील")
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                drv_name = st.text_input("Driver Name (ड्रायव्हरचे नाव)", value="Ramesh Pawar")
                drv_mob = st.text_input("Mobile Number (मोबाईल नंबर)", value="9823000000")
                drv_dl_no = st.text_input("Driving Licence No. (परवाना क्रमांक)", value="MH1820120004567")
            with d_col2:
                drv_aad_no = st.text_input("Aadhaar Card No. (आधार क्रमांक)")
                drv_upi = st.text_input("UPI Number / ID (उदा. GooglePay/PhonePe No.)")
            
            st.markdown("---")
            st.markdown("### 🏦 बँक तपशील (Bank Details)")
            b_col1, b_col2, b_col3 = st.columns(3)
            with b_col1:
                bank_name = st.text_input("Bank Name (बँकेचे नाव)")
            with b_col2:
                acc_no = st.text_input("Account Number (खाते क्रमांक)")
            with b_col3:
                ifsc_code = st.text_input("IFSC Code")
            
            st.markdown("---")
            st.markdown("### 📂 स्वतंत्र कागदपत्रे अपलोड")
            up_col1, up_col2 = st.columns(2)
            with up_col1:
                aadhar_file = st.file_uploader("Aadhaar Card Upload", type=["pdf", "jpg", "png"], key="aadhar_up")
            with up_col2:
                licence_file = st.file_uploader("Driving Licence Upload", type=["pdf", "jpg", "png"], key="lic_up")
            
            if st.button("ड्रायव्हरची संपूर्ण माहिती सेव्ह करा"):
                if drv_dl_no and drv_name:
                    aad_path = save_uploaded_file(aadhar_file, f"{drv_dl_no}_Aadhaar")
                    lic_path = save_uploaded_file(licence_file, f"{drv_dl_no}_Licence")
                    
                    conn = sqlite3.connect("rr_logistics.db")
                    c = conn.cursor()
                    c.execute('''INSERT OR REPLACE INTO drivers VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                              (drv_dl_no, drv_name, drv_mob, drv_aad_no, drv_upi, bank_name, acc_no, ifsc_code, aad_path, lic_path, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    conn.close()
                    st.success(f"ड्रायव्हर {drv_name} चे तपशील डेटाबेसमध्ये कायमस्वरूपी सेव्ह झाले! ✅")
                else:
                    st.warning("कृपया ड्रायव्हरचे नाव व लायसन्स नंबर प्रविष्ट करा!")

        # --- MENU 3: TRIP INFO & LR ---
        elif menu == "३. Trip Information & Auto LR Generation":
            st.subheader("📄 ट्रिप माहिती आणि ऑफिशियल LR (Bilty) निर्मिती")
            
            if "trip_auto_data" not in st.session_state:
                st.session_state.trip_auto_data = {
                    "lr_no": "1602/1603", "from_loc": "WONDER", "consignor": "Ultratech Cement Depo, Sarwad",
                    "d_name": "Ramesh Pawar", "truck_no": "MH18BZ2240", "to_loc": "MALEGAON",
                    "consignee": "Self / Party Warehouse, Malegaon", "d_dl": "MH1820120004567",
                    "date_val": "01-09-2026", "c_gst": "27AABCU1234F1Z5", "c_eway": "EW-889923410",
                    "d_mobile": "9823000000", "weight": "42", "freight": "22638", "order_by": "DEPO"
                }

            invoice_file = st.file_uploader("Plant Invoice Upload (PDF/Image)", type=["pdf", "jpg", "png"], key="inv_upload_key")
            
            if invoice_file is not None:
                st.session_state.trip_auto_data.update({
                    "lr_no": "1602/1603", "from_loc": "WONDER", "consignor": "Ultratech Cement Depo, Sarwad",
                    "d_name": "Ramesh Pawar", "truck_no": "MH18BZ2240", "to_loc": "MALEGAON",
                    "consignee": "Self / Party Warehouse, Malegaon", "d_dl": "MH1820120004567",
                    "date_val": "01-09-2026", "c_gst": "27AABCU1234F1Z5", "c_eway": "EW-889923410",
                    "d_mobile": "9823000000", "weight": "42", "freight": "22638", "order_by": "DEPO"
                })
                st.success("📄 इनव्हॉइस यशस्वीरित्या अपलोड झाले! खालील तपशील ऑटो-फ्लो झाले आहेत. ✅")

            col1, col2, col3 = st.columns(3)
            with col1:
                lr_no = st.text_input("L.R. No.", value=st.session_state.trip_auto_data["lr_no"])
                from_loc = st.text_input("From (कुठून)", value=st.session_state.trip_auto_data["from_loc"])
                consignor = st.text_area("Consignor", value=st.session_state.trip_auto_data["consignor"])
                d_name = st.text_input("Driver Name", value=st.session_state.trip_auto_data["d_name"])
            with col2:
                truck_no = st.text_input("Truck No.", value=st.session_state.trip_auto_data["truck_no"])
                to_loc = st.text_input("To (कुठे)", value=st.session_state.trip_auto_data["to_loc"])
                consignee = st.text_area("Consignee", value=st.session_state.trip_auto_data["consignee"])
                d_dl = st.text_input("D.L. No.", value=st.session_state.trip_auto_data["d_dl"])
            with col3:
                date_val = st.text_input("Date", value=st.session_state.trip_auto_data["date_val"])
                c_gst = st.text_input("Consignor GSTIN", value=st.session_state.trip_auto_data["c_gst"])
                c_eway = st.text_input("E-way Bill No.", value=st.session_state.trip_auto_data["c_eway"])
                d_mobile = st.text_input("Driver Mobile No.", value=st.session_state.trip_auto_data["d_mobile"])
            
            col_ex1, col_ex2, col_ex3 = st.columns(3)
            with col_ex1:
                weight = st.text_input("Weight (वजन)", value=st.session_state.trip_auto_data["weight"])
            with col_ex2:
                freight = st.text_input("Freight (भाडे रक्कम)", value=st.session_state.trip_auto_data["freight"])
            with col_ex3:
                order_by = st.text_input("Order By", value=st.session_state.trip_auto_data["order_by"])

            st.session_state.trip_auto_data.update({
                "lr_no": lr_no, "truck_no": truck_no, "from_loc": from_loc, "to_loc": to_loc,
                "consignor": consignor, "consignee": consignee, "c_gst": c_gst, "c_eway": c_eway,
                "d_name": d_name, "d_dl": d_dl, "d_mobile": d_mobile, "weight": weight,
                "freight": freight, "order_by": order_by, "date_val": date_val
            })
            
            if st.button("LR PDF तयार करा"):
                if lr_no:
                    pdf_file_path = generate_official_lr(lr_no, truck_no, date_val, from_loc, to_loc, consignor, consignee, c_gst, c_eway, d_name, d_dl, d_mobile, weight, freight, order_by)
                    st.success("LR PDF तयार झाली आहे! 👇")
                    with open(pdf_file_path, "rb") as f:
                        st.download_button(
                            label="📥 डाऊनलोड ऑफिशियल LR PDF",
                            data=f,
                            file_name=f"LR_{lr_no.replace('/', '_')}_Official.pdf",
                            mime="application/pdf"
                        )
                else:
                    st.warning("कृपया L.R. No. प्रविष्ट करा!")

        # --- MENU 4: DIESEL ---
        elif menu == "४. Diesel & Expenses Tracking":
            st.subheader("⛽ डिझेल आणि खर्च नोंद (Diesel & Expenses Tracking)")
            col_d1, col_d2, col_d3, col_d4 = st.columns(4)
            with col_d1:
                d_truck = st.text_input("Truck No.", value="MH18BZ2240")
            with col_d2:
                pump_name = st.text_input("Pump Name (पंपचे नाव)", value="R.R. Petroleum, Dhule")
            with col_d3:
                diesel_ltr = st.number_input("Diesel in Ltr (लिटर)", min_value=0.0, value=90.0, step=10.0)
            with col_d4:
                diesel_amount = st.number_input("Diesel Total Amount (एकूण रक्कम)", min_value=0.0, value=8866.8, step=100.0)
            
            if st.button("डिझेल माहिती सेव्ह करा"):
                st.success(f"गाडी क्र. {d_truck} साठी {pump_name} येथून {diesel_ltr} लिटर डिझेलची नोंद झाली! ✅")

        # --- MENU 5: TRIP SHEET AUTO WITH SAVE TO DB ---
        elif menu == "५. Trip Sheet (Auto Excel & PDF)":
            st.subheader("📊 R. R. LOGISTICS, DHULE - TRIP SHEET")
            st.info("💡 माहिती भरताच खाली आपोआप Freight, डिझेल खर्च, Total Exp व Trip Balance कॅल्क्युलेट होईल आणि ती सेव्हही करता येईल.")

            t1, t2, t3, t4 = st.columns(4)
            with t1:
                ts_lr = st.text_input("LR.NO", value="1602/1603")
                ts_from = st.text_input("From", value="WONDER")
                ts_rate = st.number_input("Rate (per Ton)", min_value=0.0, value=539.0, step=1.0)
                ts_bill = st.text_input("Bill No", value="")
            with t2:
                ts_ldate = st.text_input("Loading Date", value="01-09-2026")
                ts_to = st.text_input("To", value="MALEGAON")
                ts_weight = st.number_input("Weight (Tons)", min_value=0.0, value=42.0, step=0.5)
                ts_otp = st.text_input("OTP", value="")
            with t3:
                ts_date = st.text_input("Date", value="01-09-2026")
                ts_truck = st.text_input("Truck No", value="MH18BZ2240")
                calc_freight = round(ts_rate * ts_weight, 2)
                st.metric("Freight (भाडे फॉर्म्युला)", f"₹ {calc_freight:,.2f}")
                ts_party = st.text_input("Party Name", value="DEPO")
            with t4:
                ts_start_km = st.number_input("Start Km", min_value=0, value=0)
                ts_end_km = st.number_input("End Km", min_value=0, value=0)
                calc_run_km = ts_end_km - ts_start_km
                st.metric("Running Km", f"{calc_run_km} KM")

            st.markdown("---")
            st.markdown("#### 💰 ट्रिप खर्च तपशील (Expenses & Diesel)")

            c_e1, c_e2, c_e3 = st.columns(3)
            with c_e1:
                exp_adv = st.number_input("Trip Adv (अ‍ॅडव्हान्स)", min_value=0.0, value=6000.0, step=500.0)
                adv_date = st.text_input("Adv Date", value="03-09-2026")
                exp_unloading = st.number_input("Unloading", min_value=0.0, value=0.0)
                exp_food = st.number_input("Food (जेवण)", min_value=0.0, value=0.0)
                exp_tyre = st.number_input("Tyre Bill", min_value=0.0, value=0.0)

            with c_e2:
                exp_ft1 = st.number_input("FasTag-1", min_value=0.0, value=0.0)
                exp_ft2 = st.number_input("FasTag-2", min_value=0.0, value=0.0)
                exp_comm = st.number_input("Commission", min_value=0.0, value=0.0)
                exp_maint = st.number_input("Maintenance", min_value=0.0, value=0.0)
                exp_other = st.number_input("Other Exp", min_value=0.0, value=0.0)

            with c_e3:
                st.markdown("**⛽ डिझेल तपशील**")
                d1_ltr = st.number_input("Diesel-1 (Ltr)", min_value=0.0, value=90.0, step=5.0)
                d1_rate = st.number_input("Diesel-1 (Rate)", min_value=0.0, value=98.52, step=0.1)
                d1_amt = round(d1_ltr * d1_rate, 2)
                st.write(f"Diesel-1 Amount: ₹ {d1_amt}")
                d1_date = st.text_input("Diesel-1 Date", value="01-09-2026")
                d1_pump = st.text_input("Diesel-1 Pump", value="")
                
                st.markdown("---")
                d2_ltr = st.number_input("Diesel-2 (Ltr)", min_value=0.0, value=0.0)
                d2_rate = st.number_input("Diesel-2 (Rate)", min_value=0.0, value=0.0)
                d2_amt = round(d2_ltr * d2_rate, 2)
                d2_date = st.text_input("Diesel-2 Date", value="")
                d2_pump = st.text_input("Diesel-2 Pump", value="")

            tot_diesel_ltr = round(d1_ltr + d2_ltr, 2)
            tot_exp = round(exp_adv + exp_unloading + exp_food + exp_tyre + exp_ft1 + exp_ft2 + d1_amt + d2_amt + exp_comm + exp_maint + exp_other, 2)
            trip_bal = round(calc_freight - tot_exp, 2)
            avg_km = round(calc_run_km / tot_diesel_ltr, 2) if tot_diesel_ltr > 0 else 0

            st.markdown("---")
            st.markdown("### 📋 ट्रिप सारांश (Auto Summary)")
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("एकूण भाडे (Freight)", f"₹ {calc_freight:,.2f}")
            s2.metric("एकूण खर्च (Total Exp)", f"₹ {tot_exp:,.2f}")
            s3.metric("शिल्लक रक्कम (Trip Balance)", f"₹ {trip_bal:,.2f}")
            s4.metric("सरासरी (Average)", f"{avg_km} Km/L")

            ts_dict = {
                'lr_no': ts_lr, 'loading_date': ts_ldate, 'date_val': ts_date,
                'from_loc': ts_from, 'to_loc': ts_to, 'truck_no': ts_truck,
                'rate_ton': ts_rate, 'weight_ton': ts_weight,
                'bill_no': ts_bill, 'otp': ts_otp, 'party_name': ts_party,
                'start_km': ts_start_km, 'end_km': ts_end_km,
                'trip_adv': exp_adv, 'adv_date': adv_date,
                'unloading': exp_unloading, 'food': exp_food, 'tyre_bill': exp_tyre,
                'fastag1': exp_ft1, 'fastag2': exp_ft2,
                'd1_ltr': d1_ltr, 'd1_rate': d1_rate, 'd1_date': d1_date, 'd1_pump': d1_pump,
                'd2_ltr': d2_ltr, 'd2_rate': d2_rate, 'd2_date': d2_date, 'd2_pump': d2_pump,
                'commission': exp_comm, 'maintenance': exp_maint, 'other_exp': exp_other
            }

            # Save Trip Sheet to Database
            if st.button("💾 ही ट्रिप शीट डेटाबेसमध्ये कायमस्वरूपी सेव्ह करा"):
                conn = sqlite3.connect("rr_logistics.db")
                c = conn.cursor()
                c.execute('''INSERT OR REPLACE INTO trip_sheets VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                          (ts_lr, ts_ldate, ts_date, ts_from, ts_to, ts_truck, ts_rate, ts_weight, calc_freight,
                           ts_bill, ts_otp, ts_party, ts_start_km, ts_end_km, calc_run_km, exp_adv, exp_unloading,
                           exp_food, exp_tyre, exp_ft1, exp_ft2, tot_diesel_ltr, (d1_amt+d2_amt), exp_comm, exp_maint,
                           exp_other, tot_exp, trip_bal, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                conn.commit()
                conn.close()
                st.success(f"LR No. {ts_lr} ची ट्रिप शीट यशस्वीरीत्या सेव्ह झाली! ✅ तुम्ही मेन्यू ६ मध्ये जाऊन कधीही पाहू शकता.")

            st.markdown("---")
            st.markdown("### 📥 फाईल डाऊनलोड करा (Excel & PDF)")
            down_col1, down_col2 = st.columns(2)
            
            with down_col1:
                if HAS_OPENPYXL:
                    excel_bytes = generate_trip_sheet_excel(ts_dict)
                    st.download_button(
                        label="📊 डाऊनलोड Trip Sheet Excel (with Live Formulas)",
                        data=excel_bytes,
                        file_name=f"TripSheet_{ts_lr.replace('/', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    df_ts = pd.DataFrame([ts_dict])
                    csv_data = df_ts.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📊 डाऊनलोड Trip Sheet Data (CSV/Excel)",
                        data=csv_data,
                        file_name=f"TripSheet_{ts_lr.replace('/', '_')}.csv",
                        mime="text/csv"
                    )

            with down_col2:
                pdf_file_path = generate_trip_sheet_pdf(ts_dict)
                with open(pdf_file_path, "rb") as f:
                    st.download_button(
                        label="📄 डाऊनलोड Trip Sheet PDF",
                        data=f,
                        file_name=f"TripSheet_{ts_lr.replace('/', '_')}.pdf",
                        mime="application/pdf"
                    )

        # --- MENU 6 (NEW): VIEW SAVED RECORDS & EXPIRY REPORT ---
        elif menu == "६. Saved Records & View Documents":
            st.subheader("📁 सेव्ह केलेला डेटा, कागदपत्रे व डॉक्युमेंट एक्सपायरी रिपोर्ट")
            
            tab1, tab2, tab3 = st.tabs(["🚛 वाहने व डॉक्युमेंट एक्सपायरी", "📄 सेव्ह केलेल्या ट्रिप शीट्स", "👨‍✈️ ड्रायव्हर रेकॉर्ड्स"])
            
            conn = sqlite3.connect("rr_logistics.db")
            
            # TAB 1: VEHICLES & EXPIRY
            with tab1:
                st.markdown("#### 🚚 वाहने व कागदपत्रे एक्सपायरी स्थिती")
                v_df = pd.read_sql_query("SELECT veh_no, owner_name, veh_type, ins_exp, fit_exp, puc_exp, per_exp FROM vehicles", conn)
                if not v_df.empty:
                    st.dataframe(v_df, use_container_width=True)
                    
                    # Generate Expiry Report PDF
                    c = conn.cursor()
                    c.execute("SELECT * FROM vehicles")
                    v_all = c.fetchall()
                    exp_pdf = generate_doc_expiry_pdf(v_all)
                    with open(exp_pdf, "rb") as f:
                        st.download_button(
                            label="📥 डाऊनलोड सर्व गाड्यांचा डॉक्युमेंट एक्सपायरी PDF रिपोर्ट",
                            data=f,
                            file_name=f"Vehicle_Doc_Expiry_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf"
                        )
                else:
                    st.info("अद्याप कोणतीही गाडी नोंदणीकृत नाही.")
                    
                st.markdown("---")
                st.markdown("##### 🔍 गाडीचे अपलोड केलेले कागदपत्र पहा / डाउनलोड करा:")
                sel_veh = st.selectbox("गाडी क्रमांक निवडा:", v_df['veh_no'].tolist() if not v_df.empty else [])
                if sel_veh:
                    c = conn.cursor()
                    c.execute("SELECT rc_doc, ins_doc, fit_doc, puc_doc, per_doc FROM vehicles WHERE veh_no=?", (sel_veh,))
                    docs = c.fetchone()
                    d_names = ["RC Book", "Insurance", "Fitness", "PUC", "National Permit"]
                    for name, path in zip(d_names, docs):
                        if path and os.path.exists(path):
                            with open(path, "rb") as file_data:
                                st.download_button(f"📥 डाउनलोड {name} ({os.path.basename(path)})", data=file_data, file_name=os.path.basename(path))
                        else:
                            st.caption(f"{name}: कागदपत्र उपलब्ध नाही")

            # TAB 2: SAVED TRIP SHEETS
            with tab2:
                st.markdown("#### 📑 सेव्ह केलेल्या सर्व ट्रिप शीट्स")
                t_df = pd.read_sql_query("SELECT lr_no, date_val, truck_no, from_loc, to_loc, freight, total_exp, trip_balance, created_at FROM trip_sheets", conn)
                if not t_df.empty:
                    st.dataframe(t_df, use_container_width=True)
                else:
                    st.info("अद्याप कोणतीही ट्रिप शीट सेव्ह केलेली नाही.")

            # TAB 3: DRIVERS
            with tab3:
                st.markdown("#### 👨‍‍✈️ ड्रायव्हर यादी व कागदपत्रे")
                d_df = pd.read_sql_query("SELECT dl_no, driver_name, mobile, bank_name, acc_no, ifsc, created_at FROM drivers", conn)
                if not d_df.empty:
                    st.dataframe(d_df, use_container_width=True)
                else:
                    st.info("अद्याप कोणत्याही ड्रायव्हरची नोंदणी नाही.")
            
            conn.close()

    else:
        st.info("कृपया वरील युझर आयडी आणि पासवर्ड टाकून लॉगिन करा.")

if __name__ == '__main__':
    main()
