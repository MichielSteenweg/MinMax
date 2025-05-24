import streamlit as st
import pandas as pd
import numpy as np

def bereken_minmax(row):
    servicegraden = {
        'A': 0.99,
        'B': 0.985,
        'C': 0.98,
        'D': 0.95,
        'E': 0.95,
        'F': 0.95,
        'G': 0.95  # fallback
    }

    orderkosten = 0.50
    voorraadkosten_pct = 0.12  # per jaar (1% per maand)

    werkdagen_per_maand = 21.75
    dagverkoop = row['#6mnd'] / (6 * werkdagen_per_maand) if row['#6mnd'] > 0 else 0

    trend = 1.0
    if row['#12mnd'] > 0:
        trend = max(trend, (row['#6mnd'] * 2) / row['#12mnd'])

    abc = str(row['ABC']).strip().upper()
    serviceniveau = servicegraden.get(abc, 0.95)

    levertijd = row['Levert.'] if row['Levert.'] > 0 else 5
    cycle_waarde = row['Cyclus']
    cycle_dagen = cycle_waarde * 5 if cycle_waarde > 0 else 5
    totale_horizon = levertijd + cycle_dagen

    verwacht_gebruik = dagverkoop * totale_horizon * trend
    veiligheidsvoorraad = serviceniveau * dagverkoop * np.sqrt(totale_horizon)

    min_nieuw = veiligheidsvoorraad + verwacht_gebruik
    max_nieuw = min_nieuw + verwacht_gebruik

    kostprijs = row['Kostprijs'] / row['Per'] if row['Per'] > 0 else 0
    jaarverbruik = row['#6mnd'] * 2

    q_optimaal = np.sqrt((2 * orderkosten * jaarverbruik) / (voorraadkosten_pct * kostprijs)) if kostprijs > 0 else 0

    min_afgerond = int(np.ceil(min_nieuw))
    max_afgerond = int(np.ceil(max_nieuw))

    return pd.Series({
        'Dagverkoop': dagverkoop,
        'Trend': trend,
        'Serviceniveau': serviceniveau,
        'Min_Nieuw': min_afgerond,
        'Max_Nieuw': max_afgerond,
        'Q_optimaal': round(q_optimaal, 2)
    })

st.title("Min-Max Berekening op Basis van Verkoopdata")

uploaded_file = st.file_uploader("Upload het Excelbestand", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        resultaat = pd.concat([df, df.apply(bereken_minmax, axis=1)], axis=1)

        st.success("Berekening voltooid. Bekijk of download het resultaat hieronder.")
        st.dataframe(resultaat)

        output = pd.ExcelWriter("MinMax_Resultaat.xlsx", engine='xlsxwriter')
        resultaat.to_excel(output, index=False)
        output.close()

        with open("MinMax_Resultaat.xlsx", "rb") as f:
            st.download_button("Download Excel Resultaat", f, file_name="MinMax_Resultaat.xlsx")

    except Exception as e:
        st.error(f"Fout bij verwerken bestand: {e}")
