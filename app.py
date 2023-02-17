import streamlit as st
import datetime
from jours_feries_france import JoursFeries

from make_sheet import make_sheet

CURRENT_YEAR = datetime.datetime.now().year


st.title("Téléchargement du fichier de suivi de caisse")

year = st.selectbox(
    "Choisir une année",
    options=list(range(2020, 2100)),
    index=CURRENT_YEAR - 2020,
)


off_days = JoursFeries.for_year(year)
off_days = off_days.values()

file = make_sheet(year, off_days)

st.download_button(
    label="Télécharger le fichier",
    data=file,
    file_name=f"feuille_caisse_{year}.xlsx",
)