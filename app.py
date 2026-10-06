elif menu == "३. Trip Information & Auto LR Generation":
                st.subheader("📄 ट्रिप माहिती आणि ऑफिशियल LR (Bilty) निर्मिती")
                
                # Session State इनिशिअलाईज करणे (Auto-Flow साठी)
                if "trip_data" not in st.session_state:
                    st.session_state.trip_data = {
                        "lr_no": "", "truck_no": "", "from_loc": "Dhule", "to_loc": "AURANGABAD",
                        "consignor": "", "consignee": "", "c_gst": "", "c_eway": "",
                        "d_name": "", "d_dl": "", "d_mobile": "", "weight": "As per Plant",
                        "freight": "To Pay", "order_by": "Direct Party", "date_val": datetime.now().strftime('%d/%m/%Y')
                    }

                invoice_file = st.file_uploader("Plant Invoice Upload (PDF/Image)", type=["pdf", "jpg", "png"], key="invoice_uploader")
                
                # फाईल अपलोड झाल्यावर डेटा आपोआप सेट (Auto-Flow) करणे
                if invoice_file is not None:
                    # येथे तुम्ही मॅन्युअल किंवा फाईल डिटेक्शननुसार डिफॉल्ट व्हॅल्यू सेट करू शकता
                    st.session_state.trip_data.update({
                        "lr_no": "729",
                        "truck_no": "RJ06GC8458",
                        "from_loc": "Dhule",
                        "to_loc": "AURANGABAD",
                        "consignor": "Ultratech Cement Depo, Sarwad",
                        "consignee": "Self / Party Warehouse, Aurangabad",
                        "c_gst": "27AABCU1234F1Z5",
                        "c_eway": "EW-889923410",
                        "d_name": "Ramesh Pawar",
                        "d_dl": "MH1820120004567",
                        "d_mobile": "9823000000",
                        "weight": "As per Plant",
                        "freight": "To Pay",
                        "order_by": "Direct Party"
                    })
                    st.success("📄 इनव्हॉइस यशस्वीरित्या अपलोड झाले! खालील तपशील ऑटो-फ्लो (Auto-filled) झाले आहेत. ✅")

                # इनपुट फील्ड्स (Session State शी जोडलेले)
                col1, col2, col3 = st.columns(3)
                with col1:
                    lr_no = st.text_input("L.R. No.", value=st.session_state.trip_data["lr_no"])
                    from_loc = st.text_input("From (कुठून)", value=st.session_state.trip_data["from_loc"])
                    consignor = st.text_area("Consignor", value=st.session_state.trip_data["consignor"])
                    d_name = st.text_input("Driver Name", value=st.session_state.trip_data["d_name"])
                with col2:
                    truck_no = st.text_input("Truck No.", value=st.session_state.trip_data["truck_no"])
                    to_loc = st.text_input("To (कुठे)", value=st.session_state.trip_data["to_loc"])
                    consignee = st.text_area("Consignee", value=st.session_state.trip_data["consignee"])
                    d_dl = st.text_input("D.L. No.", value=st.session_state.trip_data["d_dl"])
                with col3:
                    date_val = st.text_input("Date", value=st.session_state.trip_data["date_val"])
                    c_gst = st.text_input("Consignor GSTIN", value=st.session_state.trip_data["c_gst"])
                    c_eway = st.text_input("E-way Bill No.", value=st.session_state.trip_data["c_eway"])
                    d_mobile = st.text_input("Driver Mobile No.", value=st.session_state.trip_data["d_mobile"])
                
                col_ex1, col_ex2, col_ex3 = st.columns(3)
                with col_ex1:
                    weight = st.text_input("Weight (वजन)", value=st.session_state.trip_data["weight"])
                with col_ex2:
                    freight = st.text_input("Freight (भाडे रक्कम)", value=st.session_state.trip_data["freight"])
                with col_ex3:
                    order_by = st.text_input("Order By", value=st.session_state.trip_data["order_by"])

                # व्हॅल्यूज अपडेट ठेवणे
                st.session_state.trip_data.update({
                    "lr_no": lr_no, "truck_no": truck_no, "from_loc": from_loc, "to_loc": to_loc,
                    "consignor": consignor, "consignee": consignee, "c_gst": c_gst, "c_eway": c_eway,
                    "d_name": d_name, "d_dl": d_dl, "d_mobile": d_mobile, "weight": weight,
                    "freight": freight, "order_by": order_by, "date_val": date_val
                })
