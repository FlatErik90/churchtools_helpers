import streamlit as st
import churchtools as ct


@st.cache_resource
def create_client():
    base_url = "https://nak.church.tools"
    client = ct.ChurchTools(base_url)
    client.login(username=st.secrets["username"], password=st.secrets["password"], remember_me=True)
    return client


st.switch_page("pages/1_Kalender_Export.py")
