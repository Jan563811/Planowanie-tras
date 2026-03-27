import streamlit as st
import pandas as pd
import requests

st.title("Geokodowanie adresów – etap 1")

# Pobranie klucza z secrets
API_KEY = st.secrets["GOOGLE_MAPS_API_KEY"]

uploaded_file = st.file_uploader("Wgraj plik XLSX z kolumną 'address'", type=["xlsx"])

def geocode_address(address):
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": address,
        "key": API_KEY
    }
    response = requests.get(url, params=params)
    data = response.json()

    if data["status"] == "OK":
        location = data["results"][0]["geometry"]["location"]
        return location["lat"], location["lng"]
    else:
        return None, None

if uploaded_file and st.button("Geokoduj"):
    df = pd.read_excel(uploaded_file)

    if "address" not in df.columns:
        st.error("Plik musi zawierać kolumnę 'address'")
    else:
        latitudes = []
        longitudes = []

        with st.spinner("Geokodowanie w toku..."):
            for address in df["address"]:
                lat, lng = geocode_address(address)
                latitudes.append(lat)
                longitudes.append(lng)

        df["latitude"] = latitudes
        df["longitude"] = longitudes

        st.success("Gotowe!")
        st.dataframe(df)

        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Pobierz wynik CSV",
            csv,
            "geocoded_results.csv",
            "text/csv"
        )