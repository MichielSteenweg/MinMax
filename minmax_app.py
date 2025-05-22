import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

# Werkdagen
def werkdagen_tussen(dagen):
    return round((dagen / 7) * 5)

def bereken_trendfactor(df):
    maandelijks_6m = df["Verkoop6M"] / 6
    maandelijks_24m = df["Verkoop24M"] / 24
    trendfactor = (maandelijks_6m / maandelijks_24m.replace(0, 0.01)).clip(lower=0.8, upper=1.2)
    return trendfactor

def bereken_dagverkoop(df):
    return df["Verkoop6M"] / 126  # 6 maanden ≈ 126 werkdagen

def bereken_optimale_bestelgrootte(df, bestelkosten, voorraadkosten_p_jaar):
    df["Jaarverbruik"] = df["Dagverkoop"] * 261
    df["EOQ"] = np.sqrt((2 * df["Jaarverbruik"] * bestelkosten) / (voorraadkosten_p_jaar * df["KostprijsPerStuk"]))
    df["EOQ_KleinerDanBestelgroote"] = df["EOQ"] < df["Bestelgroote"]
    df["OptimaleBestelgrootte"] = np.where(df["EOQ_KleinerDanBestelgroote"], df["Bestelgroote"],
                                            (df["EOQ"] / df["Bestelgroote"]).round(0) * df["Bestelgroote"])
    return df

def get_z_value_from_abc(abc):
    servicegraad_dict = {
        "A": 99.5, "B": 99.0, "C": 98.5, "D": 98.0,
        "E": 97.0, "F": 97.0, "G": 97.0
    }
    z_dict = {
        97.0: 1.88, 98.0: 2.05, 98.5: 2.17,
        99.0: 2.33, 99.5: 2.58
    }
    service = servicegraad_dict.get(str(abc).upper(), 98.0)
    z = z_dict.get(service, 2.05)
    return service, z

def bereken_min_max(df, bestelkosten, voorraadkosten_p_jaar):
    df["KostprijsPerStuk"] = df["Kostprijs"] / df["Per"].replace(0, np.nan)
    df["Dagverkoop"] = bereken_dagverkoop(df)
    df["Trendfactor"] = bereken_trendfactor(df)
    df["DagverkoopTrend"] = df["Dagverkoop"] * df["Trendfactor"]

    df = bereken_optimale_bestelgrootte(df, bestelkosten, voorraadkosten_p_jaar)

    servicegraden, z_waardes = [], []
    for abc in df["ABC"]:
        service, z = get_z_value_from_abc(abc)
        servicegraden.append(service)
        z_waardes.append(z)
    df["Servicegraad"] = servicegraden
    df["Z"] = z_waardes

    df["Dekperiode"] = df["LevertijdWD"] + df["Cyclus"] * 5
    df["Veiligheidsvoorraad"] = df["Z"] * df["DagverkoopTrend"] * np.sqrt(df["Dekperiode"])
    df["Min"] = df["DagverkoopTrend"] * df["Dekperiode"] + df["Veiligheidsvoorraad"]
    df["Max"] = df["Min"] + df["OptimaleBestelgrootte"]

    df["GemiddeldeVoorraadNieuw"] = (df["Min"] + df["Max"]) / 2
    df["GemiddeldeVoorraadHuidig"] = (df["MinHuidig"] + df["MaxHuidig"]) / 2
    df["VoorraadkostenNieuw"] = df["GemiddeldeVoorraadNieuw"] * df["KostprijsPerStuk"] * (voorraadkosten_p_jaar / 12)
    df["VoorraadkostenHuidig"] = df["GemiddeldeVoorraadHuidig"] * df["KostprijsPerStuk"] * (voorraadkosten_p_jaar / 12)
    df["VerschilVoorraadWaarde"] = (df["GemiddeldeVoorraadNieuw"] - df["GemiddeldeVoorraadHuidig"]) * df["KostprijsPerStuk"]

    return df

def genereer_excel(df):
    output = BytesIO()
    df_export_conditional = df.copy()
    decimal_minmax = ((df_export_conditional["MinHuidig"] % 1 != 0) | (df_export_conditional["MaxHuidig"] % 1 != 0))

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export_conditional.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        red_fill = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        eoql_col_letter = chr(65 + df_export_conditional.columns.get_loc("EOQ"))
        flag_col_letter = chr(65 + df_export_conditional.columns.get_loc("EOQ_KleinerDanBestelgroote"))
        aantal_rijen = len(df_export_conditional)

        for row in range(2, aantal_rijen + 2):
            worksheet.conditional_format(f"{eoql_col_letter}{row}", {
                "type": "formula",
                "criteria": f"=${flag_col_letter}{row}=TRUE",
                "format": red_fill
            })

    df_export = df_export_conditional.drop(columns=["EOQ_KleinerDanBestelgroote"]).copy()

    for col in df_export.columns:
        if col in ["Min", "Max"]:
            df_export[col] = np.where(decimal_minmax, df_export[col].round(2), df_export[col].round(0))
        elif col in df.columns[:14]:
            continue
        else:
            df_export[col] = df_export[col].round(2)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name="Sheet1")

    output.seek(0)
    return output

# Streamlit UI
st.title("Min/Max + EOQ met trend en ABC-servicegraad")

bestelkosten = 0.5
voorraadkosten_p_jaar = 0.12
abc_levels = {"A": 99.5, "B": 99.0, "C": 98.5, "D": 98.0, "E": 97.0, "F": 97.0, "G": 97.0}

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Uitlevergraad per ABC**")
    for k, v in abc_levels.items():
        st.markdown(f"{k}: {v}%")
with col2:
    st.markdown("**Parameters**")
    st.markdown(f"Voorraadkosten: **1% p/mnd** ({voorraadkosten_p_jaar * 100:.1f}% p/j)")
    st.markdown(f"Bestelkosten: **€{bestelkosten:.2f}** per orderregel")

uploaded_file = st.file_uploader("Upload Excel-bestand met artikeldata", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("Voorbeeld van ingelezen data:", df.head())

        verplichte_kolommen = {
            "Verkoop1M", "Verkoop2M", "Verkoop6M", "Verkoop12M", "Verkoop24M",
            "LevertijdWD", "Cyclus", "Kostprijs", "Per", "Bestelgroote", "Artikelnummer",
            "MinHuidig", "MaxHuidig", "ABC"
        }
        if verplichte_kolommen.issubset(df.columns):
            resultaat_df = bereken_min_max(df, bestelkosten, voorraadkosten_p_jaar)
            totaal_verschil = resultaat_df["VerschilVoorraadWaarde"].sum()
            st.metric("Totaal verschil voorraadwaarde (€)", f"{totaal_verschil:,.2f}")

            excel_bestand = genereer_excel(resultaat_df)
            st.download_button("Download resultaat als Excel", excel_bestand, file_name="minmax_resultaat.xlsx")
        else:
            st.error(f"Het bestand mist één of meer verplichte kolommen: {verplichte_kolommen}")
    except Exception as e:
        st.error(f"Fout bij verwerken van bestand: {e}")
