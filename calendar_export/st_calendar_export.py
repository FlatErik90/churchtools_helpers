import streamlit as st
import pandas as pd
import churchtools as ct
import datetime

st.set_page_config(layout="wide")
st.header("ChurchTools Kalender-Export")


@st.cache_resource
def create_client():
    client = ct.ChurchTools("https://nak.church.tools")
    client.login(username=st.secrets["username"], password=st.secrets["password"], remember_me=True)
    return client


client = create_client()

calenders = [c for c in client.calendars.list() if c.isPrivate is False]
selected_calenders = st.sidebar.multiselect(label="Kalender",
                                      placeholder="Wähle deine Kalender",
                                      options=calenders,
                                      format_func=lambda x: x.name)
days = st.sidebar.number_input(label="Zeitraum (Tage)", value=14)

show = st.sidebar.button(label="Anzeigen")

if show:
    appointments = client.calendars.appointments([c.id for c in selected_calenders],
                                                 datetime.datetime.now(),
                                                 datetime.datetime.now() + datetime.timedelta(days=days))
    fields = ['startDate', 'endDate',  'caption', 'calendar', 'information', 'note', 'allDay']
    data = [{fn: getattr(f, fn) for fn in fields} for f in appointments]
    for d in data:
        c = d["calendar"]
        d["calendar"] = c.name
        # st.write(type(d["startDate"]))
        if isinstance(d["startDate"], datetime.date):
            d["startDate"] = d["startDate"].strftime("%d.%m.%Y - %H:%M Uhr")
        if isinstance(d["endDate"], datetime.date):
            d["endDate"] = d["endDate"].strftime("%d.%m.%Y - %H:%M Uhr")
    # st.write(data)
    df = pd.DataFrame(data)
    st.dataframe(df, hide_index=True, column_config={"caption": "Bezeichnung",
                                                     "startDate": "Start",
                                                     "startTime": "Uhrzeit",
                                                     "endDate": "Ende",
                                                     "calendar": "Kalender",
                                                     "information": "Infos",
                                                     "note": "Notizen",
                                                     "allDay": "Ganztätig"})
