import streamlit as st
import pandas as pd
from io import BytesIO

def create_excel_template():
    # Define the columns
    columns = [
        "Vārds uzvārds", "Uzņēmums", "Valsts kods (XX)", "Adrese 1",
        "Adrese 2", "Pasta indekss", "Grupas", "E-pasts", "Tālrunis",
        "Sūtījuma klase", "Sūtījuma tips", "Sūtījuma veids", "Apdrošināšana, €",
        "Pēcmaksa, €", "Sūtījuma svars", "Sūtījuma saturs", "Nosaukums",
        "Daudzums", "Neto svars (kg)", "Vērtība, €", "HS tarifa Nr.",
        "Izcelsmes valsts", "Papildus pakalpojumi", "Komerciāla prece",
        "Piezīmes", "AP", "MD", "PVN Nr./ Eksportētāja kods",
        "PVN Nr./ Importētāja kods", "Postage Paid", "Saistītie dokumenti",
        "Dokumenta apraksts", "Dokumenta numurs"
    ]

    # Create an empty DataFrame with the defined columns
    df = pd.DataFrame(columns=columns)

    return df

def process_csv_data(uploaded_file):
    # Read the CSV file
    df_csv = pd.read_csv(uploaded_file)

    # Initialize the Excel template
    df_excel = create_excel_template()

    # Extract "Valsts kods (XX)" as "LV" if "Adrese" contains it
    if "Adrese" in df_csv.columns:
        df_excel["Valsts kods (XX)"] = df_csv["Adrese"].apply(lambda x: "LV" if "LV" in str(x) else "").str.strip()

    # Extract "Pasta indekss" as the part containing "LV-" and following digits
    if "Adrese" in df_csv.columns:
        df_excel["Pasta indekss"] = df_csv["Adrese"].str.extract(r'(LV-\d+)')[0]

    # Extract the last line before "LV" or similar pattern as "Adrese 1"
    if "Adrese" in df_csv.columns:
        df_excel["Adrese 1"] = df_csv["Adrese"].str.extract(r'([A-Za-zĀ-Žā-ž\s\.]+(?:nov\.|pag\.|pils\.)?)\s*,?\s*LV')[0].str.strip()

    # Extract all lines except the last one before "LV" as "Adrese 2"
    if "Adrese" in df_csv.columns:
        df_excel["Adrese 2"] = df_csv["Adrese"].str.replace(r'(,\s*LV.*)', '', regex=True).str.strip()

    # Set constant values for "Uzņēmums" and "Grupas"
    df_excel["Uzņēmums"] = "SIA METRUM"
    df_excel["Grupas"] = "Klienti 1"

    # Copy "Mērnieks_Vārds_Uzvārds" to "Vārds uzvārds"
    if "Mērnieks_Vārds_Uzvārds" in df_csv.columns:
        df_excel["Vārds uzvārds"] = df_csv["Mērnieks_Vārds_Uzvārds"]

    return df_excel

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dati')
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("Excel Veidnes Izveide un Datu Pievienošana")
    st.write("Augšupielādējiet CSV failu, lai pievienotu datus Excel veidnei.")

    uploaded_file = st.file_uploader("Augšupielādējiet CSV failu", type="csv")

    if uploaded_file is not None:
        df_excel = process_csv_data(uploaded_file)

        st.success("Dati veiksmīgi apstrādāti un pievienoti Excel veidnei!")

        # Show the resulting DataFrame
        st.write("### Rezultāta Tabula")
        st.dataframe(df_excel)

        excel_data = to_excel(df_excel)
        st.download_button(
            label="Lejupielādēt Excel failu",
            data=excel_data,
            file_name='apstrādāta_veidne.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()
