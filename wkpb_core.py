import pandas as pd
from datetime import datetime
import sys

WKPB_KOLOMMEN = [
    "identificatie",
    "monumentnummer",
    "register",
    "kenmerk",
    "kadastraleGemeenteBRK",
    "sectieBRK",
    "perceelnummerBRK"
]

ACTIEF_KOLOM = "indicatieObjectVervallenBRK"


def controleer_kolommen(df):
    vereiste = WKPB_KOLOMMEN + [ACTIEF_KOLOM]
    ontbrekend = [k for k in vereiste if k not in df.columns]

    if ontbrekend:
        print(f"Ontbrekende kolommen: {', '.join(ontbrekend)}")
        sys.exit(1)


def is_actief(waarde):
    if pd.isna(waarde):
        return False
    waarde = str(waarde).strip().upper()
    return waarde in ["WAAR", "TRUE", "1"]


def actieve_telling(df):
    actief_mask = df[ACTIEF_KOLOM].apply(is_actief)
    actief = df[actief_mask]
    telling = actief.groupby("identificatie").size()
    return telling


def main():
    df_oud = pd.read_excel("oud.xlsx")
    df_nieuw = pd.read_excel("nieuw.xlsx")

    controleer_kolommen(df_oud)
    controleer_kolommen(df_nieuw)

    oud_telling = actieve_telling(df_oud)
    nieuw_telling = actieve_telling(df_nieuw)

    actief_df = pd.DataFrame({
        "oud_actief": oud_telling,
        "nieuw_actief": nieuw_telling
    }).fillna(0)

    werklijst_ids = actief_df[
        (actief_df["oud_actief"] == 0) &
        (actief_df["nieuw_actief"] > 0)
    ].index

    werklijst = (
        df_nieuw[df_nieuw["identificatie"].isin(werklijst_ids)]
        .sort_values("identificatie")
        .drop_duplicates("identificatie", keep="first")
        .copy()
    )

    werklijst["Wijzigingstype"] = "Gewijzigd"
    werklijst = werklijst[WKPB_KOLOMMEN + ["Wijzigingstype"]]

    audit_df = actief_df.loc[werklijst_ids].reset_index()
    audit_df["reden"] = "nieuw actief BRK-object"

    datum = datetime.now().strftime("%Y%m%d")
    output = f"wkpb_werklijst_{datum}.xlsx"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        werklijst.to_excel(writer, sheet_name="WKPB Lijst", index=False)
        pd.DataFrame({
            "Type": ["Gewijzigd", "Totaal"],
            "Aantal": [len(werklijst), len(werklijst)]
        }).to_excel(writer, sheet_name="Samenvatting", index=False)
        audit_df.to_excel(writer, sheet_name="Audit verschil", index=False)


if __name__ == "__main__":
    main()