import streamlit as st
import pandas as pd
import io
import numpy as np

import datetime
import pytz
import locale
locale.setlocale(locale.LC_ALL, "de_DE.utf8")

st.set_page_config(layout="wide")


from Home import create_client
from document_utils import dump_calendar

# with st.sidebar:
#     nav = st.container(border=True)
#     nav.page_link("pages/1_Kalender_Export.py", label="Kalender Export")
#     nav.page_link("pages/2_Dienste_Export.py", label="Dienste Export")

st.header("ChurchTools Kalender-Export")


client = create_client()

timezone = pytz.timezone("Europe/Berlin")

calenders = [c for c in client.calendars.list() if not c.isPrivate and c.name != "Amtsträger"]
selected_calenders = st.sidebar.multiselect(label="Kalender",
                                      placeholder="Wähle deine Kalender",
                                      options=calenders,
                                      format_func=lambda x: x.name,
                                      default=calenders)

start_end = st.sidebar.date_input("Zeitraum (Datum)", value=(datetime.datetime.now(), datetime.datetime.now() + datetime.timedelta(days=28)))
days = st.sidebar.number_input(label="Zeitraum (Tage)", value=28)
if len(start_end) == 1:
    start = start_end[0]
    end = start + datetime.timedelta(days=days)
else:
    start = start_end[0]
    end = start_end[1]
hide_regular_services = st.sidebar.checkbox(label="Normale Gottesdienste ausblenden", value=True)
remove_duplicates = st.sidebar.checkbox(label="Doppelte Einträge ausblenden", value=True)
hide_lessons = st.sidebar.checkbox(label="Unterrichte ausblenden", value=True)

df = None
highlight_rows = None
if len(selected_calenders) > 0:
    appointments = client.calendars.appointments([c.id for c in selected_calenders], start, end)
    if hide_regular_services:
        appointments = [a for a in appointments if a.caption != "Gottesdienst" or a.note is not None]
    if hide_lessons:
        appointments = [a for a in appointments if a.caption != "Sonntagsschule" and 
                        a.caption != "Religionsunterricht" and a.caption != "Konfirmandenunterricht"]
    if len(appointments) > 0:
        # print(appointments)
        fields = ['startDate', 'endDate',  'caption', 'calendar', 'information', 'note', 'allDay', 'address']
        data = [{fn: getattr(f, fn) for fn in fields} for f in appointments]
        for d in data:
            c = d["calendar"]
            d["calendar"] = c.name
            if not d["allDay"]:
                d["startTime"] = d["startDate"].astimezone(timezone).strftime("%H:%M")
                d["endTime"] = d["endDate"].astimezone(timezone).strftime("%H:%M")
            d["weekDay"] = d["startDate"].strftime("%A")
            d["startDate"] = d["startDate"].strftime("%d. %B")
            d["endDate"] = d["endDate"].strftime("%d. %B")
            # print(str(d["startDate"]), str(d["endDate"]))
            if str(d["startDate"]) == str(d["endDate"]):
                d["endDate"] = ""
            if d["address"] is not None:
                d["place"] = d["address"].meetingAt


        df = pd.DataFrame(data)
        if remove_duplicates:
            df.drop_duplicates(subset=["startDate", "startTime", "caption"], inplace=True)
        column_map = { "weekDay": "Wochentag",
                       "startDate": "Datum",
                       "startTime": "Uhrzeit",
                       "caption": "Termin",
                       "note": "Untertitel",
                       "place": "Ort",
                        # "endDate": "Ende (Datum)",
                        # "endTime": "Ende (Uhrzeit)",
                        "calendar": "Kalender"
                        # "information": "Infos",
                        # "allDay": "Ganztätig"
                       }
        for key, value in column_map.items():
            df[value] = df.get(key, None)
            # df.drop(columns=[key])
        for col in df.columns:
            if col not in column_map.values():
                df.drop(columns=[col], inplace=True)
        # print(list(column_map.values()))
        df = df.loc[:, list(column_map.values())]
        st.dataframe(df, hide_index=True)
        highlight_rows = []
        current_date = df.iloc[0].Datum
        for i, entry in df.iterrows():
            # if i == 0:
            #     continue
            # if entry.Datum == current_date:
            #     df.at[i, "Datum"] = None
            #     df.at[i, "Wochentag"] = None
            current_date = entry.Datum
            if not isinstance(entry.Untertitel, float) and entry.Untertitel is not None:
                df.at[i, "Termin"] = df.at[i, "Termin"] + "\n" + df.at[i, "Untertitel"]
            if not isinstance(entry.Ort, float) and entry.Ort is not None:
                df.at[i, "Termin"] = df.at[i, "Termin"] + "\n" + df.at[i, "Ort"]
            if entry.Kalender.startswith("Gottesdienste"):
                highlight_rows.append(i)
        # st.write(highlight_rows)
        df.drop(labels=["Untertitel", "Ort", "Kalender"], axis="columns", inplace=True)
    else:
        st.info("Keine Termine gefunden.")

if df is not None:
    output_buffer = io.BytesIO()
    dump_calendar(df, start.strftime("%B"), highlight_rows, output_buffer)
    save_as_excel = st.download_button(label="Als Excelsheet exportieren", data=output_buffer, file_name="Kalender.xlsx")

