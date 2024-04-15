import streamlit as st
import churchtools as ct


@st.cache_resource
def create_client():
    base_url = "https://nak.church.tools"
    client = ct.ChurchTools(base_url)
    client.login(username=st.secrets["username"], password=st.secrets["password"], remember_me=True)
    return client


nav = st.container(border=True)
nav.page_link("pages/1_Kalender_Export.py", label="Kalender Export")
nav.page_link("pages/2_Dienste_Export.py", label="Dienste Export")
# st.switch_page("pages/1_Kalender_Export.py")
