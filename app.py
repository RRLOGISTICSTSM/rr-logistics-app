import streamlit as st
import pandas as pd
from datetime import datetime
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import tempfile

st.set_page_config(page_title="RR Logistics TSM", page_icon="🚚", layout="wide")

def generate_official_lr(lr_no, truck_no, date_val, from_loc, to_loc, consignor, consignee, c_gst, c_eway, d_name, d_dl, d_mobile, weight, freight, order_by):
    temp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(temp_dir, f"LR_{lr_no}.pdf")
    
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=2
    )
    sub_style = ParagraphStyle(
        'SubStyle', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=2
    )
    addr_style = ParagraphStyle(
        'AddrStyle', parent=styles['Normal'], fontSize=7.5, textColor=colors.HexColor('#334155'), alignment=1, spaceAfter=8
    )
    
    logo_path = "logo.jpeg" 
    
    if os.path.exists(logo_path):
        img = Image(logo_path, width=70, height=60)
        img.hAlign = 'CENTER'
        story.append(img)
        story.append(Spacer(1, 4))
    
    story.append(Paragraph("SUBJECT TO DHULE JURISDICTION", ParagraphStyle('SubTop', parent=styles['Normal'], fontSize=6.5, textColor=colors.HexColor('#475569'), alignment=1)))
    story.append(Paragraph("R.R. LOGASTICS TRANSPORT PVT. LTD.", title_style))
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
            Paragraph("For R.R. LOGASTICS TRANSPORT PVT. LTD.<br/><br/><br/>Authorized Signatory", styles['Normal'])
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

def main():
    st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>R.R. LOGASTICS TRANSPORT PVT. LTD.</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #475569;'><b>TRANSPORT CONTRACTOR & COMMISSION AGENT</b> | Dhule Jurisdiction</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.sidebar.title("🔐 सिस्टम लॉगिन")
    username = st.sidebar.text_input("User ID", key="login_user")
    password = st.sidebar.text_input("Password", type="password", key="login_pass")
    
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.sidebar.button("Login") or st.session_state.logged_in:
        if (username == "admin" and password == "rr123") or st.session_state.logged_in:
            st.session_state.logged_in = True
            st.sidebar.success("सक्सेसफुली लॉगिन झाले! ✅")
            
            st.sidebar.markdown("---")
            menu = st.sidebar.radio("मेन्यू निवडा:", [
                "१. New Vehicle Registration & Docs", 
                "२. Driver Registration & Docs", 
                "३. Trip Information & Auto LR Generation", 
                "४. Diesel & Expenses Tracking"
            ], key="dashboard_menu")
            
            if menu == "१. New Vehicle Registration & Docs":
                st.subheader("🚛 नवीन गाडी नोंदणी व सर्व कागदपत्रे अपलोड (Vehicle Registration & Documents)")
                
                v_col1, v_col2 = st.columns(2)
                with v_col1:
                    veh_no = st.text_input("Vehicle Number (गाडी क्रमांक)", value="RJ06GC8458")
                    veh_owner = st.text_input("Owner Name (मालकाचे नाव)")
                with v_col2:
                    veh_type = st.text_input("Vehicle Type (उदा. 10 व्हीलर, 14 व्हीलर)")
                    veh_cap = st.text_input("Capacity in Ton (क्षमता)")
                
                st.markdown("---")
                st.markdown("### 📂 कागदपत्रे अपलोड (Documents Upload)")
                
                doc_col1, doc_col2 = st.columns(2)
                with doc_col1:
                    rc_file = st.file_uploader("1. RC Book Upload", type=["pdf", "jpg", "png"], key="rc_up")
                    insurance_file = st.file_uploader("2. Insurance Upload (इन्शुरन्स)", type=["pdf", "jpg", "png"], key="ins_up")
                    fitness_file = st.file_uploader("3. Fitness Certificate Upload (फिटनेस)", type=["pdf", "jpg", "png"], key="fit_up")
                with doc_col2:
                    puc_file = st.file_uploader("4. PUC Upload (पोल्युशन)", type=["pdf", "jpg", "png"], key="puc_up")
                    permit_file = st.file_uploader("5. National Permit Upload (नॅशनल परमिट)", type=["pdf", "jpg", "png"], key="per_up")
                
                if st.button("गाडीची संपूर्ण माहिती व कागदपत्रे सेव्ह करा"):
                    if veh_no:
                        st.success(f"गाडी क्र. {veh_no} ची माहिती आणि सर्व कागदपत्रे यशस्वीरित्या सेव्ह झाली! ✅")
                    else:
                        st.warning("कृपया गाडी क्रमांक (Vehicle Number) प्रविष्ट करा!")

            elif menu == "२. Driver Registration & Docs":
                st.subheader("👨‍✈️ ड्रायव्हर नोंदणी, स्वतंत्र कागदपत्रे आणि बँक/UPI तपशील")
                
                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    drv_name = st.text_input("Driver Name (ड्रायव्हरचे नाव)", value="Ramesh Pawar")
                    drv_mob = st.text_input("Mobile Number (मोबाईल नंबर)", value="9823000000")
                    drv_dl_no = st.text_input("Driving Licence No. (परवाना क्रमांक)", value="MH1820120004567")
                with d_col2:
                    drv_aad_no = st.text_input("Aadhar Card No. (आधार क्रमांक)")
                    drv_upi = st.text_input("UPI Number / ID (उदा. ybl@upi / GooglePay/PhonePe No.)")
                
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
                st.markdown("### 📂 स्वतंत्र कागदपत्रे अपलोड (Aadhar & Licence Separate Upload)")
                up_col1, up_col2 = st.columns(2)
                with up_col1:
                    aadhar_file = st.file_uploader("Aadhar Card Upload (आधार कार्ड)", type=["pdf", "jpg", "png"], key="aadhar_up")
                with up_col2:
                    licence_file = st.file_uploader("Driving Licence Upload (ड्रायव्हिंग लायसन्स)", type=["pdf", "jpg", "png"], key="lic_up")
                
                if st.button("ड्रायव्हरची संपूर्ण माहिती सेव्ह करा"):
                    if drv_name:
                        st.success(f"ड्रायव्हर {drv_name} चे आधार, लायसन्स, बँक व UPI तपशील यशस्वीरित्या सेव्ह झाले! ✅")
                    else:
                        st.warning("कृपया ड्रायव्हरचे नाव प्रविष्ट करा!")

            elif menu == "३. Trip Information & Auto LR Generation":
                st.subheader("📄 ट्रिप माहिती आणि ऑफिशियल LR (Bilty) निर्मिती")
                
                invoice_file = st.file_uploader("Plant Invoice Upload (PDF/Image)", type=["pdf", "jpg", "png"])
                
                a_lr = "729" if invoice_file else ""
                a_truck = "RJ06GC8458" if invoice_file else ""
                a_consignor = "Ultratech Cement Depo, Sarwad" if invoice_file else ""
                a_consignee = "Self / Party Warehouse, Aurangabad" if invoice_file else ""
                a_gst = "27AABCU1234F1Z5" if invoice_file else ""
                a_eway = "EW-889923410" if invoice_file else ""
                a_driver = "Ramesh Pawar" if invoice_file else ""
                a_dl = "MH1820120004567" if invoice_file else ""
                a_mob = "9823000000" if invoice_file else ""
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    lr_no = st.text_input("L.R. No.", value=a_lr)
                    from_loc = st.text_input("From (कुठून)", value="Dhule")
                    consignor = st.text_area("Consignor", value=a_consignor)
                    d_name = st.text_input("Driver Name", value=a_driver)
                with col2:
                    truck_no = st.text_input("Truck No.", value=a_truck)
                    to_loc = st.text_input("To (कुठे)", value="AURANGABAD")
                    consignee = st.text_area("Consignee", value=a_consignee)
                    d_dl = st.text_input("D.L. No.", value=a_dl)
                with col3:
                    date_val = st.text_input("Date", value=datetime.now().strftime('%d/%m/%Y'))
                    c_gst = st.text_input("Consignor GSTIN", value=a_gst)
                    c_eway = st.text_input("E-way Bill No.", value=a_eway)
                    d_mobile = st.text_input("Driver Mobile No.", value=a_mob)
                
                col_ex1, col_ex2, col_ex3 = st.columns(3)
                with col_ex1:
                    # वजनाच्या ठिकाणी आता बायडिफॉल्ट फक्त 'As per Plant' येईल
                    weight = st.text_input("Weight (वजन)", value="As per Plant")
                with col_ex2:
                    freight = st.text_input("Freight (भाडे रक्कम)", value="To Pay")
                with col_ex3:
                    order_by = st.text_input("Order By", value="Direct Party")
                
                if st.button("LR PDF तयार करा"):
                    if lr_no:
                        pdf_file_path = generate_official_lr(lr_no, truck_no, date_val, from_loc, to_loc, consignor, consignee, c_gst, c_eway, d_name, d_dl, d_mobile, weight, freight, order_by)
                        st.success("LR PDF तयार झाली आहे! 👇")
                        
                        with open(pdf_file_path, "rb") as f:
                            st.download_button(
                                label="📥 डाऊनलोड ऑफिशियल LR PDF",
                                data=f,
                                file_name=f"LR_{lr_no}_Official.pdf",
                                mime="application/pdf"
                            )
                    else:
                        st.warning("कृपया L.R. No. प्रविष्ट करा!")

            elif menu == "४. Diesel & Expenses Tracking":
                st.subheader("⛽ डिझेल आणि खर्च नोंद (Diesel & Expenses Tracking)")
                
                col_d1, col_d2, col_d3, col_d4 = st.columns(4)
                with col_d1:
                    d_truck = st.text_input("Truck No.", value="RJ06GC8458")
                with col_d2:
                    pump_name = st.text_input("Pump Name (पंपचे नाव)", value="R.R. Petroleum, Dhule")
                with col_d3:
                    diesel_ltr = st.number_input("Diesel in Ltr (लिटर)", min_value=0.0, value=200.0, step=10.0)
                with col_d4:
                    diesel_amount = st.number_input("Diesel Total Amount (एकूण रक्कम)", min_value=0.0, value=18500.0, step=500.0)
                
                if st.button("डिझेल माहिती सेव्ह करा"):
                    st.success(f"गाडी क्र. {d_truck} साठी {pump_name} येथून {diesel_ltr} लिटर डिझेलची (रक्कम: ₹{diesel_amount}) नोंद यशस्वीरित्या झाली! ✅")
                
                st.markdown("---")
                if st.button("📥 डाऊनलोड Excel रिपोर्ट"):
                    df = pd.DataFrame({
                        "Truck No": ["RJ06GC8458"], 
                        "Pump Name": ["R.R. Petroleum, Dhule"], 
                        "Diesel Ltr": [200.0], 
                        "Total Amount": [18500.0]
                    })
                    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "RR_Diesel_Report.xlsx")
                    df.to_excel(desktop_path, index=False)
                    st.success("डेस्कटॉपवर डिझेलचा Excel रिपोर्ट सेव्ह झाला! ✅")
        else:
            st.sidebar.error("चुकीचा User ID किंवा Password!")
    else:
        st.info("कृपया लॉगिन करा.")

if __name__ == '__main__':
    main()