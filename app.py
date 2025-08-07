import streamlit as st
import pandas as pd
import io
import csv
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Table de correspondance Diametre -> Lettre pour FP2E
diametre_lettre = {
    15: ['A', 'U', 'V'],
    20: ['B'],
    25: ['C'],
    30: ['D'],
    40: ['E'],
    50: ['F'],
    60: ['G'],
    65: ['G'],
    80: ['H'],
    100: ['I'],
    125: ['J'],
    150: ['K']
}

def get_csv_delimiter(file):
    """
    Détecte automatiquement le délimiteur d'un fichier CSV.
    """
    try:
        sample = file.read(2048).decode('utf-8')
        dialect = csv.Sniffer().sniff(sample)
        file.seek(0)
        return dialect.delimiter
    except Exception:
        file.seek(0)
        return ','

def check_data(df):
    """
    Vérifie les données du DataFrame pour détecter les anomalies en utilisant des opérations vectorisées.
    Retourne un DataFrame avec les lignes contenant des anomalies.
    """
    df_with_anomalies = df.copy()

    # Vérification des colonnes requises
    required_columns = ['Protocole Radio', 'Marque', 'Numéro de compteur', 'Numéro de tête', 'Latitude', 'Longitude', 'Année de fabrication', 'Diametre', 'Traité']
    if not all(col in df_with_anomalies.columns for col in required_columns):
        missing = [col for col in required_columns if col not in df_with_anomalies.columns]
        st.error(f"Colonnes requises manquantes : {', '.join(missing)}")
        st.stop()

    df_with_anomalies['Anomalie'] = ''

    # Conversion des colonnes pour les analyses et remplacement des NaN par des chaînes vides
    df_with_anomalies['Numéro de compteur'] = df_with_anomalies['Numéro de compteur'].astype(str).fillna('')
    df_with_anomalies['Numéro de tête'] = df_with_anomalies['Numéro de tête'].astype(str).fillna('')
    df_with_anomalies['Marque'] = df_with_anomalies['Marque'].astype(str).fillna('')
    df_with_anomalies['Protocole Radio'] = df_with_anomalies['Protocole Radio'].astype(str).fillna('')
    df_with_anomalies['Traité'] = df_with_anomalies['Traité'].astype(str).fillna('')

    # Marqueurs pour les conditions
    is_kamstrup = df_with_anomalies['Marque'].str.upper() == 'KAMSTRUP'
    is_sappel = df_with_anomalies['Marque'].str.upper().isin(['SAPPEL (C)', 'SAPPEL (H)', 'SAPPEL(C)'])
    is_itron = df_with_anomalies['Marque'].str.upper() == 'ITRON'

    annee_fabrication_num = pd.to_numeric(df_with_anomalies['Année de fabrication'], errors='coerce')
    df_with_anomalies['Diametre'] = pd.to_numeric(df_with_anomalies['Diametre'], errors='coerce')

    # ------------------------------------------------------------------
    # ANOMALIES GÉNÉRALES (valeurs manquantes et incohérences de base)
    # ------------------------------------------------------------------
    
    # Colonnes manquantes
    df_with_anomalies.loc[df_with_anomalies['Protocole Radio'].isin(['', 'nan']), 'Anomalie'] += 'Protocole Radio manquant / '
    df_with_anomalies.loc[df_with_anomalies['Marque'].isin(['', 'nan']), 'Anomalie'] += 'Marque manquante / '
    df_with_anomalies.loc[df_with_anomalies['Numéro de compteur'].isin(['', 'nan']), 'Anomalie'] += 'Numéro de compteur manquant / '
    df_with_anomalies.loc[df_with_anomalies['Diametre'].isnull(), 'Anomalie'] += 'Diamètre manquant / '
    df_with_anomalies.loc[annee_fabrication_num.isnull(), 'Anomalie'] += 'Année de fabrication manquante / '
    
    # Numéro de tête manquant (sauf pour Kamstrup)
    condition_tete_manquante = (df_with_anomalies['Numéro de tête'].isin(['', 'nan'])) & (~is_kamstrup)
    df_with_anomalies.loc[condition_tete_manquante, 'Anomalie'] += 'Numéro de tête manquant / '

    # Coordonnées
    df_with_anomalies.loc[df_with_anomalies['Latitude'].isnull() | df_with_anomalies['Longitude'].isnull(), 'Anomalie'] += 'Coordonnées GPS non numériques / '
    coord_invalid = ((df_with_anomalies['Latitude'] == 0) | (~df_with_anomalies['Latitude'].between(-90, 90))) | \
                    ((df_with_anomalies['Longitude'] == 0) | (~df_with_anomalies['Longitude'].between(-180, 180)))
    df_with_anomalies.loc[coord_invalid, 'Anomalie'] += 'Coordonnées GPS invalides / '

    # ------------------------------------------------------------------
    # ANOMALIES SPÉCIFIQUES AUX MARQUES
    # ------------------------------------------------------------------
    
    # KAMSTRUP
    kamstrup_valid = is_kamstrup & (~df_with_anomalies['Numéro de tête'].isin(['', 'nan']))
    df_with_anomalies.loc[is_kamstrup & (df_with_anomalies['Numéro de compteur'].str.len() != 8), 'Anomalie'] += 'KAMSTRUP: Compteur ≠ 8 caractères / '
    df_with_anomalies.loc[kamstrup_valid & (df_with_anomalies['Numéro de compteur'] != df_with_anomalies['Numéro de tête']), 'Anomalie'] += 'KAMSTRUP: Compteur ≠ Tête / '
    df_with_anomalies.loc[kamstrup_valid & (~df_with_anomalies['Numéro de compteur'].str.isdigit() | ~df_with_anomalies['Numéro de tête'].str.isdigit()), 'Anomalie'] += 'KAMSTRUP: Compteur ou Tête non numérique / '
    
    # SAPPEL
    sappel_valid = is_sappel & (~df_with_anomalies['Numéro de tête'].isin(['', 'nan']))
    df_with_anomalies.loc[sappel_valid & (df_with_anomalies['Numéro de tête'].str.len() != 16), 'Anomalie'] += 'SAPPEL: Tête ≠ 16 caractères / '
    df_with_anomalies.loc[is_sappel & (~df_with_anomalies['Numéro de compteur'].str.match(r'^[A-Z]{1}\d{2}[A-Z]{2}\d{6}$')), 'Anomalie'] += 'SAPPEL: Compteur format incorrect / '
    df_with_anomalies.loc[(is_sappel) & (df_with_anomalies['Numéro de compteur'].str.startswith('C', na=False)) & (df_with_anomalies['Marque'].str.upper() != 'SAPPEL (C)'), 'Anomalie'] += 'SAPPEL: Incohérence Marque/Compteur (C) / '
    df_with_anomalies.loc[(is_sappel) & (df_with_anomalies['Numéro de compteur'].str.startswith('H', na=False)) & (df_with_anomalies['Marque'].str.upper() != 'SAPPEL (H)'), 'Anomalie'] += 'SAPPEL: Incohérence Marque/Compteur (H) / '
    
    # ITRON
    itron_valid = is_itron & (~df_with_anomalies['Numéro de tête'].isin(['', 'nan']))
    df_with_anomalies.loc[itron_valid & (df_with_anomalies['Numéro de tête'].str.len() != 8), 'Anomalie'] += 'ITRON: Tête ≠ 8 caractères / '
    df_with_anomalies.loc[is_itron & (~df_with_anomalies['Numéro de compteur'].str.lower().str.startswith(('i', 'd'), na=False)), 'Anomalie'] += 'ITRON: Compteur doit commencer par "I" ou "D" / '

    # ------------------------------------------------------------------
    # ANOMALIES MULTI-COLONNES (logique consolidée)
    # ------------------------------------------------------------------

    # Protocole Radio vs Traité
    traite_lra_condition = df_with_anomalies['Traité'].str.startswith(('903', '863'), na=False)
    condition_radio_lra = traite_lra_condition & (df_with_anomalies['Protocole Radio'].str.upper() != 'LRA')
    df_with_anomalies.loc[condition_radio_lra, 'Anomalie'] += 'Protocole ≠ LRA pour Traité 903/863 / '
    
    condition_radio_sgx = (~traite_lra_condition) & (df_with_anomalies['Protocole Radio'].str.upper() != 'SGX')
    df_with_anomalies.loc[condition_radio_sgx, 'Anomalie'] += 'Protocole ≠ SGX pour Traité non 903/863 / '

    # Règle de diamètre FP2E (pour SAPPEL)
    sappel_fp2e_condition = (is_sappel) & (df_with_anomalies['Numéro de compteur'].str.len() >= 5) & (df_with_anomalies['Diametre'].notnull())
    def check_fp2e(row):
        compteur = row['Numéro de compteur']
        annee_compteur = compteur[1:3]
        annee_fabrication = str(int(row['Année de fabrication'])) if pd.notna(row['Année de fabrication']) else ''
        if len(annee_fabrication) < 2 or annee_compteur != annee_fabrication[-2:]:
            return False
        
        lettre_diam = compteur[4].upper()
        return lettre_diam in diametre_lettre.get(row['Diametre'], [])

    fp2e_anomalies = df_with_anomalies[sappel_fp2e_condition].apply(lambda row: not check_fp2e(row), axis=1)
    df_with_anomalies.loc[sappel_fp2e_condition & fp2e_anomalies, 'Anomalie'] += 'SAPPEL: non conforme FP2E / '

    # Nettoyage de la colonne 'Anomalie'
    df_with_anomalies['Anomalie'] = df_with_anomalies['Anomalie'].str.strip().str.rstrip(' /')
    
    anomalies_df = df_with_anomalies[df_with_anomalies['Anomalie'] != ''].copy()
    anomalies_df.reset_index(inplace=True)
    anomalies_df.rename(columns={'index': 'Index original'}, inplace=True)
    
    # Comptage des anomalies pour le résumé
    anomaly_counter = anomalies_df['Anomalie'].str.split(' / ').explode().value_counts()
    
    return anomalies_df, anomaly_counter

def afficher_resume_anomalies(anomaly_counter):
    """
    Affiche un résumé des anomalies.
    """
    if not anomaly_counter.empty:
        summary_df = pd.DataFrame(anomaly_counter).reset_index()
        summary_df.columns = ["Type d'anomalie", "Nombre de cas"]
        st.subheader("Récapitulatif des anomalies")
        st.dataframe(summary_df)

# --- Interface Streamlit ---
st.title("Contrôle des données de Télérelève")
st.markdown("Veuillez téléverser votre fichier pour lancer les contrôles.")

uploaded_file = st.file_uploader("Choisissez un fichier", type=['csv', 'xlsx'])

if uploaded_file is not None:
    st.success("Fichier chargé avec succès !")

    try:
        file_extension = uploaded_file.name.split('.')[-1]
        if file_extension == 'csv':
            delimiter = get_csv_delimiter(uploaded_file)
            df = pd.read_csv(uploaded_file, sep=delimiter)
        elif file_extension == 'xlsx':
            df = pd.read_excel(uploaded_file)
        else:
            st.error("Format de fichier non pris en charge. Veuillez utiliser un fichier .csv ou .xlsx.")
            st.stop()
    except Exception as e:
        st.error(f"Erreur de lecture du fichier : {e}")
        st.stop()

    st.subheader("Aperçu des 5 premières lignes")
    st.dataframe(df.head())

    if st.button("Lancer les contrôles"):
        st.write("Contrôles en cours...")
        anomalies_df, anomaly_counter = check_data(df)

        if not anomalies_df.empty:
            st.error("Anomalies détectées !")
            st.dataframe(anomalies_df)
            afficher_resume_anomalies(anomaly_counter)
            
            # Dictionnaire pour mapper les anomalies aux colonnes
            anomaly_columns_map = {
                "Protocole Radio manquant": ['Protocole Radio'],
                "Marque manquante": ['Marque'],
                "Numéro de compteur manquant": ['Numéro de compteur'],
                "Numéro de tête manquant": ['Numéro de tête'],
                "Coordonnées GPS non numériques": ['Latitude', 'Longitude'],
                "Coordonnées GPS invalides": ['Latitude', 'Longitude'],
                "Diamètre manquant": ['Diametre'],
                "Année de fabrication manquante": ['Année de fabrication'],
                
                # Anomalies spécifiques
                "KAMSTRUP: Compteur ≠ Tête": ['Numéro de compteur', 'Numéro de tête'],
                "KAMSTRUP: Compteur ou Tête non numérique": ['Numéro de compteur', 'Numéro de tête'],
                "SAPPEL: Tête ≠ 16 caractères": ['Numéro de tête'],
                "SAPPEL: Compteur format incorrect": ['Numéro de compteur'],
                "SAPPEL: Incohérence Marque/Compteur (C)": ['Marque', 'Numéro de compteur'],
                "SAPPEL: Incohérence Marque/Compteur (H)": ['Marque', 'Numéro de compteur'],
                "ITRON: Tête ≠ 8 caractères": ['Numéro de tête'],
                "ITRON: Compteur doit commencer par \"I\" ou \"D\"": ['Numéro de compteur'],
                "Protocole ≠ LRA pour Traité 903/863": ['Protocole Radio', 'Traité'],
                "Protocole ≠ SGX pour Traité non 903/863": ['Protocole Radio', 'Traité'],
                "SAPPEL: non conforme FP2E": ['Numéro de compteur', 'Diametre', 'Année de fabrication'],
            }

            if file_extension == 'csv':
                csv_file = anomalies_df.to_csv(index=False, sep=delimiter).encode('utf-8')
                st.download_button(
                    label="Télécharger les anomalies en CSV",
                    data=csv_file,
                    file_name='anomalies_telerelève.csv',
                    mime='text/csv',
                )
            elif file_extension == 'xlsx':
                excel_buffer = io.BytesIO()
                
                # Création d'un classeur Excel avec le DataFrame d'anomalies
                wb = Workbook()
                ws_anomalies = wb.active
                ws_anomalies.title = "Anomalies"
                
                # Écriture du DataFrame des anomalies dans la première feuille
                for r_idx, row in enumerate(dataframe_to_rows(anomalies_df, index=False, header=True)):
                    ws_anomalies.append(row)

                # Création d'une nouvelle feuille pour le résumé
                ws_resume = wb.create_sheet(title="Résumé des anomalies")

                # Création du DataFrame de résumé pour l'exportation
                if not anomaly_counter.empty:
                    summary_df = pd.DataFrame(anomaly_counter).reset_index()
                    summary_df.columns = ["Type d'anomalie", "Nombre de cas"]

                    # Ajout des données de résumé à la deuxième feuille
                    for r_idx, row in enumerate(dataframe_to_rows(summary_df, index=False, header=True)):
                        ws_resume.append(row)
                        
                    # Dictionnaire pour trouver l'index de la première anomalie
                    first_anomaly_index = {
                        anomaly_type: anomalies_df.index[anomalies_df['Anomalie'].str.contains(anomaly_type, regex=False)].min()
                        for anomaly_type in summary_df["Type d'anomalie"]
                    }

                    # Mise en forme de la feuille de résumé et ajout des liens
                    for cell in ws_resume['A']:
                        if cell.value in first_anomaly_index:
                            target_row = anomalies_df.index.get_loc(first_anomaly_index[cell.value]) + 2
                            cell.hyperlink = f"#'Anomalies'!A{target_row}"
                            cell.style = "Hyperlink"
                    
                # Mise en forme de la feuille des anomalies et mise en couleur
                red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

                for i, row in enumerate(anomalies_df.iterrows()):
                    anomalies = str(row[1]['Anomalie']).split(' / ')
                    for anomaly in anomalies:
                        anomaly_key = anomaly.strip()
                        if anomaly_key in anomaly_columns_map:
                            columns_to_highlight = anomaly_columns_map[anomaly_key]
                            for col_name in columns_to_highlight:
                                try:
                                    col_index = list(anomalies_df.columns).index(col_name) + 1
                                    cell = ws_anomalies.cell(row=i + 2, column=col_index)
                                    cell.fill = red_fill
                                except ValueError:
                                    pass

                # Enregistrement du classeur dans le buffer
                excel_buffer_styled = io.BytesIO()
                wb.save(excel_buffer_styled)
                excel_buffer_styled.seek(0)

                st.download_button(
                    label="Télécharger les anomalies en Excel",
                    data=excel_buffer_styled,
                    file_name='anomalies_telerelève.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
        else:
            st.success("Aucune anomalie détectée ! Les données sont conformes.")
