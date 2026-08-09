import hmac

import streamlit as st

st.title(":material/lock: Login")

credentials = st.secrets.get("credentials", {})

if not credentials:
    st.error(
        "No login credentials configured. Add a `[credentials]` section with "
        "`username` and `password` to `.streamlit/secrets.toml`."
    )
    st.stop()

_, center, _ = st.columns([1, 2, 1])

with center:
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", width="stretch")

    if submitted:
        valid = (
            hmac.compare_digest(username, str(credentials.get("username", "")))
            and hmac.compare_digest(password, str(credentials.get("password", "")))
        )
        if valid:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.rerun()
        else:
            st.error("Invalid username or password.")
