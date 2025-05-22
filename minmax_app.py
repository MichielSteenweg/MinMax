import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np


def genereer_excel(df):
    output = BytesIO()
    df_export = df.copy()
    decimal_minmax = ((df_export["MinHuidig"] % 1 != 0) | (df_export["MaxHuidig"] % 1 != 0))

    for col in df_export.columns:
        if col in ["Min", "Max"]:
            df_export[col] = np.where(decimal_minmax, df_export[col].round(2), df_export[col].round(0))
        elif col in df.columns[:14]:
            continue
        else:
            df_export[col] = df_export[col].round(2)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        red_fill = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        eoql_col_letter = chr(65 + df_export.columns.get_loc("EOQ"))
        flag_col_letter = chr(65 + df_export.columns.get_loc("EOQ_KleinerDanBestelgroote"))
        max_col_letter = chr(65 + df_export.columns.get_loc("Max"))
        aantal_rijen = len(df_export)

        for row in range(2, aantal_rijen + 2):
            worksheet.conditional_format(f"{eoql_col_letter}{row}", {
                "type": "formula",
                "criteria": f"=${flag_col_letter}{row}=TRUE",
                "format": red_fill
            })
            worksheet.conditional_format(f"{max_col_letter}{row}", {
                "type": "formula",
                "criteria": f"=${flag_col_letter}{row}=TRUE",
                "format": red_fill
            })

        # Verberg de vlagkolom
        flag_col_idx = df_export.columns.get_loc("EOQ_KleinerDanBestelgroote")
        worksheet.set_column(flag_col_idx, flag_col_idx, None, None, {"hidden": True})

        # Tweede tabblad: samenvatting met aangepaste kolomnamen
        overzicht_kolommen = {
            "Artikelnummer": "Art.",
            "Omschrijving": "Omschr.",
            "KostprijsPerStuk": "Kostprijs",
            "ABC": "ABC",
            "Courantie": "Cour.",
            "Verkoop6M": "6mnd",
            "Verkoop12M": "12mnd",
            "Verkoop24M": "24mnd",
            "Bestelgroote": "Besteleenh.",
            "LevertijdWD": "Levertijd",
            "Cyclus": "Cyclus",
            "Dekperiode": "Levert.tot.",
            "Trendfactor": "Trendf.",
            "MinHuidig": "Min",
            "MaxHuidig": "Max",
            "EOQ": "EOQ",
            "OptimaleBestelgrootte": "OBG",
            "Min": "MinN",
            "Max": "MaxN",
            "VerschilVoorraadWaarde": "Delta"
        }

        kolom_volgorde = list(overzicht_kolommen.keys())
        samenvatting_df = df_export[kolom_volgorde].copy()
        samenvatting_df.rename(columns=overzicht_kolommen, inplace=True)

        samenvatting_df.to_excel(writer, sheet_name="Overzicht", index=False)

    output.seek(0)
    return output

