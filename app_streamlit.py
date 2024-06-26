import pandas as pd
import numpy as np
import streamlit as st
import io
import openpyxl

# Titre de l'application
st.title("Comparaison soldes clients Compta / Sextan")

# Section d'importation des fichiers Excel
st.header("Importer les fichiers excel")
compta_file = st.file_uploader("Importer le fichier compta ici", type = ["xlsx"])
sextan_file = st.file_uploader("Importer le fichier sextan ici", type = ["xlsx"])

if compta_file and sextan_file:
    # Lire les fichiers Excel téléchargés
    df_compta = pd.read_excel(compta_file)
    df_sextan = pd.read_excel(sextan_file)

    # Affichage d'un message de confirmation de téléchargement
    st.success("Fichiers importés avec succès !")

    totalTTC = df_sextan["Total TTC"]
    reglement = df_sextan["Réglèment Total"]

    df_sextan["Solde_sextan"] = totalTTC - reglement

    df_sextan["Nº Client"].fillna(0, inplace=True)
    df_sextan["Nº Client"] = df_sextan["Nº Client"].astype(int)

    df_compta["Numéro"].fillna(0, inplace=True)
    df_compta["Numéro"] = df_compta["Numéro"].astype(int)

    def normalize_numero(numero) :
        numero_normalized = numero.replace("900", "")
        numero_normalized = numero_normalized.lstrip("0")
        return numero_normalized
    df_compta["Numéro"] = df_compta["Numéro"].astype(str).apply(normalize_numero)

    df_compta["Numéro"] = df_compta["Numéro"].astype(str)
    df_sextan["Nº Client"] = df_sextan["Nº Client"].astype(str)

    merged_df = pd.merge(df_compta, df_sextan, left_on = "Numéro", right_on = "Nº Client", how = "left")
    merged_df = merged_df.reset_index(drop = True)
    # df_resultats = merged_df.drop_duplicates(subset = ["N° Pjt"])
    df_resultats = merged_df.dropna(subset = ["Numéro"])
    df_resultats = df_resultats[["Numéro", "Client", "Solde", "Solde_sextan"]]

    df_resultats["Solde"] = pd.to_numeric(df_resultats["Solde"], errors = "coerce")
    df_resultats["Solde_sextan"] = pd.to_numeric(df_resultats["Solde_sextan"], errors = "coerce")

    df_resultats_grouped = df_resultats.groupby("Numéro").agg({
        "Client" : "first",
        "Solde" : "first",
        "Solde_sextan" : "sum"
    }).reset_index()

    solde = df_resultats_grouped["Solde"]
    solde_sextan = df_resultats_grouped["Solde_sextan"]
    df_resultats_grouped["Ecart"] = solde - solde_sextan

    df_resultats_grouped = df_resultats_grouped.rename(columns = { "Solde" : "Solde_compta"})

    sum_solde_compta = df_resultats_grouped["Solde_compta"].sum()
    sum_solde_sextan = df_resultats_grouped["Solde_sextan"].sum()
    sum_ecart = df_resultats_grouped["Ecart"].sum()

    sum_row = {"Numéro" : "Total", "Client" : "", "Solde_compta" : sum_solde_compta, "Solde_sextan" : sum_solde_sextan, "Ecart" : sum_ecart}
    # df_resultats_grouped = df_resultats_grouped.concat(sum_row, ignore_index = True)
    df_resultats_grouped = pd.concat([df_resultats_grouped, pd.DataFrame([sum_row])], ignore_index = True)

    # Affichage du résultat :
    st.header("Clients communs et leurs soldes")
    st.dataframe(df_resultats_grouped)

    # Fonction de conversion du df en .xlsx
    @st.cache_data
    def convert_df(df) :
        # Création d'un objet BytesIO pour stocker les données excel en mémoire
        excel_data = io.BytesIO()
        # Utilisation de to_excel pour écrire le df dans l'objet BytesIO
        df.to_excel(excel_data, index = False)
        return excel_data

    # Conversion des résultats sous forme de fichier Excel :
    excel_data = convert_df(df_resultats_grouped)

    # Bouton de téléchargement des résultats sous format .xlsx
    st.download_button(
        label = "Télécharger les résultat au format Excel",
        data = excel_data,
        file_name = "Soldes_compta_sextan.xlsx",
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else :
    st.info("Maeva , faut mettre les 2 excels pour que ça fonctionne ^^")