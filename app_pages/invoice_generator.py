import pandas as pd
import streamlit as st

import invoice_core as core

with st.sidebar:
    st.caption(f"Signed in as **{st.session_state.get('username', '')}**")
    if st.button("Log out", icon=":material/logout:", width="stretch"):
        st.session_state.authenticated = False
        st.session_state.pop("username", None)
        st.rerun()

st.title(":material/receipt_long: Invoice Generator")

s_no = st.text_input(
    "S No (required, unique 6-7 digit number)",
    max_chars=7,
    help="Mandatory. If this S No already has an invoice, the existing invoice is returned for download instead of a new one being created.",
)
name = st.text_input("Bill To (Customer Name)")
course = st.selectbox("Select Course", core.COURSES)
amount = st.number_input("Amount", min_value=0, step=100)
date = st.date_input("Receipt Date")

# --------------------------
# GENERATE
# --------------------------

if st.button("Generate Invoice"):
    if not core.is_valid_sno(s_no):
        st.error("S No is required: enter a unique 6-7 digit number.")
    else:
        df = pd.read_excel(core.EXCEL_FILE)
        existing = core.find_by_sno(df, s_no)

        if existing is not None:
            st.info(f"S No {existing['S No']} already exists as invoice {existing['Invoice No']} — showing the existing invoice.")
            data = existing
            pdf = core.ensure_pdf(data)
        elif not name or amount == 0:
            st.error("Fill all fields")
            data = None
        else:
            data = {
                "S No": s_no.strip(),
                "Invoice No": core.generate_invoice_no(df),
                "UID": core.generate_uid(df),
                "Date": str(date),
                "Name": name,
                "Course": course,
                "Amount": int(amount),
            }

            # Save to Excel
            df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
            df.to_excel(core.EXCEL_FILE, index=False)

            # Generate PDF
            pdf = core.generate_pdf(data)

            st.success(f"Invoice {data['Invoice No']} created")

        if data is not None:
            st.write("### Preview")
            st.json(data)

            # Download
            with open(pdf, "rb") as f:
                st.download_button(
                    "Download Invoice",
                    f,
                    file_name=f"{data['Invoice No']}.pdf",
                    key=f"gen_dl_{data['Invoice No']}",
                )

# --------------------------
# SEARCH
# --------------------------

st.markdown("---")
st.subheader(":material/search: Search Invoices")

search = st.text_input("Search by Name / Invoice No / S No")

if search:
    df = pd.read_excel(core.EXCEL_FILE)
    result = df[
        df["Name"].astype(str).str.contains(search, case=False, na=False, regex=False) |
        df["Invoice No"].astype(str).str.contains(search, case=False, na=False, regex=False) |
        df["S No"].apply(core.clean_id).str.contains(search, case=False, na=False, regex=False)
    ]
    st.dataframe(result, width="stretch")

    if result.empty:
        st.caption("No matching invoices.")
    else:
        for _, row in result.iterrows():
            data = row.to_dict()
            data["S No"] = core.clean_id(data.get("S No"))
            data["UID"] = core.clean_id(data.get("UID"))
            invoice_no = str(data["Invoice No"])

            with st.container(border=True, horizontal=True, vertical_alignment="center"):
                st.write(
                    f"**{invoice_no}**  \n"
                    f"S No: {data['S No'] or '—'} · {data['Name']} · {data['Course']}"
                )
                pdf_path = core.ensure_pdf(data)
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "Download",
                        f,
                        file_name=f"{invoice_no}.pdf",
                        key=f"search_dl_{invoice_no}",
                    )
