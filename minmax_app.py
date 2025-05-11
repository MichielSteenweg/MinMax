import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np
import matplotlib.pyplot as plt

# Functie om werkelijke werkdagen (excl. weekend) te berekenen
def werkdagen_tussen(dagen):
    return round((dagen / 7) * 5)

def bereken_dagverkoop(df):
    return df["Verkoop2M"] / 42  # 2 maanden ≈ 42 werkdagen

def bereken_optimale_bestelgrootte(df):
    bestelkosten = 1  # euro per orderregel
    voorraadkosten_p_jaar = 0.12  # 1% per maand = 12% per jaar

    df["Jaarverbruik"] = df["Dagverkoop"] * 261
    df["EOQ"] = np.sqrt((2 * df["Jaarverbruik"] * bestelkosten) / (voorraadkosten_p_jaar * df["Kostprijs"]))
    df["OptimaleBestelgrootte"] = (df["EOQ"] / df["Bestelgroote"]).round(0) * df["Bestelgroote"]

    return df

def get_z_value(service_level):
    z_table = {
        90: 1.28,
        95: 1.65,
        98: 2.05,
        99: 2.33,
        99.9: 3.08
    }
    return z_table.get(service_level, 2.33)  # default is 99%

def bereken_min_max(df, service_level):
    df["Dagverkoop"] = bereken_dagverkoop(df)
    df = bereken_optimale_bestelgrootte(df)

    z = get_z_value(service_level)
    df["Dekperiode"] = df["LevertijdWD"] + df["Cyclus"] * 5
    df["Veiligheidsvoorraad"] = z * df["Dagverkoop"] * np.sqrt(df["Dekperiode"])
    df["Min"] = df["Dekperiode"] * df["Dagverkoop"] + df["Veiligheidsvoorraad"]
    df["Max"] = df["Min"] + df["OptimaleBestelgrootte"]

    df["GemiddeldeVoorraadNieuw"] = (df["Min"] + df["Max"]) / 2 * df["Kostprijs"]
    df["GemiddeldeVoorraadHuidig"] = (df["MinHuidig"] + df["MaxHuidig"]) / 2 * df["Kostprijs"]
    df["VerschilVoorraadWaarde"] = df["GemiddeldeVoorraadNieuw"] - df["GemiddeldeVoorraadHuidig"]

    return df

def genereer_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

def plot_verschil(df):
    fig, ax = plt.subplots(figsize=(10, 5))
    df_sorted = df.sort_values("VerschilVoorraadWaarde", ascending=False).head(20)
    ax.bar(df_sorted["Artikelnummer"], df_sorted["VerschilVoorraadWaarde"])
    ax.set_ylabel("Verschil voorraadwaarde (€)")
    ax.set_title("Top 20 artikelen met grootste voorraadverschil")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    return fig

st.title("Min/Max + EOQ Berekening met instelbare Servicegraad")

uploaded_file = st.file_uploader("Upload Excel-bestand met artikeldata", type=["xlsx"])

service_level = st.selectbox("Kies gewenste uitlevergraad (%)", options=[90, 95, 98, 99, 99.9], index=3)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("Voorbeeld van ingelezen data:", df.head())

        verplichte_kolommen = {"Verkoop2M", "LevertijdWD", "Cyclus", "Kostprijs", "Bestelgroote", "Artikelnummer", "MinHuidig", "MaxHuidig"}
        if verplichte_kolommen.issubset(df.columns):
            resultaat_df = bereken_min_max(df, service_level)
            st.write(f"Resultaat met {service_level}% servicegraad:", resultaat_df)

            totaal_verschil = resultaat_df["VerschilVoorraadWaarde"].sum()
            st.metric("Verschil in gemiddelde voorraadwaarde (€)", f"{totaal_verschil:,.2f}")

            st.pyplot(plot_verschil(resultaat_df))

            excel_bestand = genereer_excel(resultaat_df)
            st.download_button("Download resultaat als Excel", excel_bestand, file_name="minmax_resultaat.xlsx")
        else:
            st.error(f"Het bestand mist één of meer verplichte kolommen: {verplichte_kolommen}")
    except Exception as e:
        st.error(f"Fout bij verwerken van bestand: {e}")