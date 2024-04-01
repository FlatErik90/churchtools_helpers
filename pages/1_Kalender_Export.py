import streamlit as st
import pandas as pd

import datetime
import pytz

from Home import create_client
from document_utils import dump_calendar

st.set_page_config(layout="wide")
st.header("ChurchTools Kalender-Export")


client = create_client()

timezone = pytz.timezone("Europe/Berlin")

calenders = [c for c in client.calendars.list() if not c.isPrivate and c.name != "Amtsträger"]
selected_calenders = st.sidebar.multiselect(label="Kalender",
                                      placeholder="Wähle deine Kalender",
                                      options=calenders,
                                      format_func=lambda x: x.name)
days = st.sidebar.number_input(label="Zeitraum (Tage)", value=14)
hide_regular_services = st.sidebar.checkbox(label="Normale Gottesdienste ausblenden", value=True)
remove_duplicates = st.sidebar.checkbox(label="Doppelte Einträge ausblenden", value=True)

df = None
if len(selected_calenders) > 0:
    appointments = client.calendars.appointments([c.id for c in selected_calenders],
                                                 datetime.datetime.now(), #- datetime.timedelta(days=days),
                                                 datetime.datetime.now() + datetime.timedelta(days=days))
    if hide_regular_services:
        appointments = [a for a in appointments if a.caption != "Gottesdienst" or a.note is not None]
    if len(appointments) > 0:
        fields = ['startDate', 'endDate',  'caption', 'calendar', 'information', 'note', 'allDay']
        data = [{fn: getattr(f, fn) for fn in fields} for f in appointments]
        for d in data:
            c = d["calendar"]
            d["calendar"] = c.name
            if not d["allDay"]:
                d["startTime"] = d["startDate"].astimezone(timezone).strftime("%H:%M Uhr")
                d["endTime"] = d["endDate"].astimezone(timezone).strftime("%H:%M Uhr")
            d["startDate"] = d["startDate"].strftime("%d.%m.%Y")
            d["endDate"] = d["endDate"].strftime("%d.%m.%Y")
            # print(str(d["startDate"]), str(d["endDate"]))
            if str(d["startDate"]) == str(d["endDate"]):
                d["endDate"] = ""

        df = pd.DataFrame(data)
        if remove_duplicates:
            df.drop_duplicates(subset=["startDate", "startTime", "caption"], inplace=True)
        st.dataframe(df,
                     hide_index=True,
                     column_config={"caption": "Bezeichnung",
                                    "ort": "Untertitel",
                                    "startDate": "Start (Datum)",
                                    "startTime": "Start (Uhrzeit)",
                                    "endDate": "Ende (Datum)",
                                    "endTime": "Ende (Uhrzeit)",
                                    "calendar": "Kalender",
                                    "information": "Infos",
                                    "note": "Notizen",
                                    "allDay": "Ganztätig"},
                     column_order=["startDate",
                                   "startTime",
                                   "endDate",
                                   "endTime",
                                   "caption",
                                   "ort",
                                   "calendar",
                                   "information",
                                   "note",
                                   "allDay"])

    else:
        st.info("Keine Termine gefunden.")

save_as_excel = st.button(label="Excel exportieren")
if save_as_excel and df is not None:
    dump_calendar(df)
