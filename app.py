import streamlit as st

st.set_page_config(page_title="MyCaptain Invoice", page_icon=":material/receipt_long:")

st.session_state.setdefault("authenticated", False)

login_page = st.Page("app_pages/login.py", title="Login", icon=":material/lock:")
invoice_page = st.Page(
    "app_pages/invoice_generator.py", title="Invoice generator", icon=":material/receipt_long:"
)

pages = [invoice_page] if st.session_state.authenticated else [login_page]

nav = st.navigation(pages, position="hidden")
nav.run()
