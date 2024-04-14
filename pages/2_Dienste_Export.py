import streamlit as st
import pytz
import datetime
import pandas as pd
import io

import locale
locale.setlocale(locale.LC_ALL, "de_DE.utf8")

st.set_page_config(layout="wide")

from Home import create_client
from document_utils import dump_services

with st.sidebar:
    nav = st.container(border=True)
    nav.page_link("pages/1_Kalender_Export.py", label="Kalender Export")
    nav.page_link("pages/2_Dienste_Export.py", label="Dienste Export")

client = create_client()


st.header("ChurchTools Dienste-Export")

timezone = pytz.timezone("Europe/Berlin")

days = st.sidebar.number_input(label="Zeitraum (Tage)", value=28)
events = [e for e in client.events.list(datetime.datetime.now(), #- datetime.timedelta(days=days),
                                         datetime.datetime.now() + datetime.timedelta(days=days))]


client = create_client()

service_map = {}
for service in client.services.list():
    service_map[service.id] = service
# st.write(service_map)

fields = ['id', 'startDate']
data = [{fn: getattr(f, fn) for fn in fields} for f in events]


df = None
if len(events) > 0:
    for e in data:
        e["Datum"] = e["startDate"].strftime("%d.%m")
        eventServices = client.events.get(e['id']).eventServices
        #st.write(e.eventServices)
        for s in eventServices:
            service_type = service_map[s.serviceId].name
            #st.write(service_type)
            e[service_type] = s.name
        e.pop("startDate")
        e.pop("id")
        # st.write(e)

    df = pd.DataFrame(data)
    df = df.loc[:, ["Datum", "Predigt", "Co-Predigt", "Bibellesung", "OD 1", "OD 2", "Chorleitung", "Orgel",
                    "Telefongottesdienst"]]
    st.dataframe(df, hide_index=True)

else:
    st.info("Keine Dienste gefunden.")


if df is not None:
    output_buffer = io.BytesIO()
    dump_services(df, output_buffer)
    save_as_excel = st.download_button(label="Als Excelsheet exportieren", data=output_buffer, file_name="Dienste.xlsx")
