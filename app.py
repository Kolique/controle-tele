import streamlit as st
import pandas as pd
import io
import csv
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

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

    # Conversion des colonnes pour les analyses
    df_with_anomalies['Numéro de compteur'] = df_with_anomalies['Numéro de compteur'].astype(str)
    df_with_anomalies['Numéro de tête'] = df_with_anomalies['Numéro de tête'].astype(str)
    
    is_kamstrup = df_with_anomalies['Marque'].str.upper() == 'KAMSTRUP'
    is_sappel = df_with_anomalies['Marque'].str.upper().isin(['SAPPEL (C)', 'SAPPEL (H)', 'SAPPEL(C)'])
    is_itron = df_with_anomalies['Marque'].str.upper() == 'ITRON'

    annee_fabrication_num = pd.to_numeric(df_with_anomalies['Année de fabrication'], errors='coerce')
    
    # ------------------------------------------------------------------
    # ANOMALIES GÉNÉRALES
    # ------------------------------------------------------------------
    # Colonnes manquantes
    df_with_anomalies.loc[df_with_anomalies['Protocole Radio'].isnull(), 'Anomalie'] += 'Protocole Radio manquant / '
    df_with_anomalies.loc[df_with_anomalies['Marque'].isnull(), 'Anomalie'] += 'Marque manquante / '
    df_with_anomalies.loc[df_with_anomalies['Numéro de compteur'].str.lower() == 'nan', 'Anomalie'] += 'Numéro de compteur manquant / '
    
    # Numéro de tête manquant (sauf pour Kamstrup)
    condition_tete_manquante = (df_with_anomalies['Numéro de tête'].str.lower() == 'nan') & (~is_kamstrup)
    df_with_anomalies.loc[condition_tete_manquante, 'Anomalie'] += 'Numéro de tête manquant / '

    # Coordonnées
    df_with_anomalies['Latitude'] = pd.to_numeric(df_with_anomalies['Latitude'], errors='coerce')
    df_with_anomalies['Longitude'] = pd.to_numeric(df_with_anomalies['Longitude'], errors='coerce')
    coord_invalid = ((df_with_anomalies['Latitude'] == 0) | (~df_with_anomalies['Latitude'].between(-90, 90))) | \
                    ((df_with_anomalies['Longitude'] == 0) | (~df_with_anomalies['Longitude'].between(-180, 180)))
    df_with_anomalies.loc[coord_invalid, 'Anomalie'] += 'Coordonnées invalides / '

    # Diamètre
    df_with_anomalies['Diametre'] = pd.to_numeric(df_with_anomalies['Diametre'], errors='coerce')
    df_with_anomalies.loc[df_with_anomalies['Diametre'].isnull(), 'Anomalie'] += 'Diamètre manquant / '

    # ------------------------------------------------------------------
    # ANOMALIES SPÉCIFIQUES AUX MARQUES
    # ------------------------------------------------------------------
    
    # KAMSTRUP
    kamstrup_condition = is_kamstrup & (df_with_anomalies['Numéro de tête'].str.lower() != 'nan')
    df_with_anomalies.loc[kamstrup_condition & (df_with_anomalies['Numéro de compteur'] != df_with_anomalies['Numéro de tête']), 'Anomalie'] += 'KAMSTRUP: Compteur ≠ Tête / '
    df_with_anomalies.loc[kamstrup_condition & (~df_with_anomalies['Numéro de compteur'].str.isdigit() | ~df_with_anomalies['Numéro de tête'].str.isdigit()), 'Anomalie'] += 'KAMSTRUP: Compteur ou Tête non numérique / '

    # SAPPEL
    sappel_condition_len_tete = is_sappel & (df_with_anomalies['Numéro de tête'].str.lower() != 'nan') & (df_with_anomalies['Numéro de tête'].str.len() != 16)
    df_with_anomalies.loc[sappel_condition_len_tete, 'Anomalie'] += 'SAPPEL: Tête ≠ 16 caractères / '
    
    sappel_condition_compteur = is_sappel & (~df_with_anomalies['Numéro de compteur'].str.match(r'^[A-Z]{1}\d{2}[A-Z]{2}\d{6}$'))
    df_with_anomalies.loc[sappel_condition_compteur, 'Anomalie'] += 'SAPPEL: Compteur format incorrect / '

    # ITRON
    itron_condition_len_tete = is_itron & (df_with_anomalies['Numéro de tête'].str.lower() != 'nan') & (df_with_anomalies['Numéro de tête'].str.len() != 8)
    df_with_anomalies.loc[itron_condition_len_tete, 'Anomalie'] += 'ITRON: Tête ≠ 8 caractères / '
    itron_condition_compteur = is_itron & (~df_with_anomalies['Numéro de compteur'].str.lower().str.startswith(('i', 'd')))
    df_with_anomalies.loc[itron_condition_compteur, 'Anomalie'] += 'ITRON: Compteur doit commencer par "I" ou "D" / '


    # RÈGLES BASÉES SUR PLUSIEURS COLONNES
    
    # Marque vs Numéro de compteur
    condition_marque_compteur_c = is_sappel & (df_with_anomalies['Numéro de compteur'].str.startswith('C')) & (df_with_anomalies['Marque'].str.upper() != 'SAPPEL (C)')
    df_with_anomalies.loc[condition_marque_compteur_c, 'Anomalie'] += 'SAPPEL: Incohérence Marque/Compteur (C) / '
    
    condition_marque_compteur_h = is_sappel & (df_with_anomalies['Numéro de compteur'].str.startswith('H')) & (df_with_anomalies['Marque'].str.upper() != 'SAPPEL (H)')
    df_with_anomalies.loc[condition_marque_compteur_h, 'Anomalie'] += 'SAPPEL: Incohérence Marque/Compteur (H) / '

    # Protocole Radio vs Traité
    # J'ai ajouté .fillna(False) pour gérer les NaN avant d'appliquer l'opérateur ~
    traite_lra = df_with_anomalies['Traité'].str.startswith(('903', '863'), na=False)
    traite_sgx = ~traite_lra
    
    condition_radio_lra = traite_lra & (df_with_anomalies['Protocole Radio'].str.upper() != 'LRA')
    df_with_anomalies.loc[condition_radio_lra, 'Anomalie'] += 'Protocole ≠ LRA pour Traité 903/863 / '
    
    condition_radio_sgx = traite_sgx & (df_with_anomalies['Protocole Radio'].str.upper() != 'SGX')
    df_with_anomalies.loc[condition_radio_sgx, 'Anomalie'] += 'Protocole ≠ SGX pour Traité non 903/863 / '

    # Règle de diamètre FP2E (pour SAPPEL, ITRON)
    # L'année dans le compteur est l'année à deux chiffres, ex: C23...
    def check_fp2e_vectorized(row):
        try:
            compteur = row['Numéro de compteur']
            annee_compteur = compteur[1:3]
            annee_fabrication = str(int(row['Année de fabrication'])).zfill(2)[-2:]
            lettre_diam = compteur[4]
            diametre = int(row['Diametre'])
            
            if annee_compteur != annee_fabrication:
                return 'Année dans compteur ≠ Année de fabrication'
            
            lettres_attendues = diametre_lettre.get(diametre, [])
            if lettre_diam not in lettres_attendues:
                return 'Lettre de diamètre incohérente'
                
            return None
        except:
            return 'Règle FP2E non applicable'

    fp2e_anomalies = df_with_anomalies[is_sappel | is_itron].apply(check_fp2e_vectorized, axis=1)
    df_with_anomalies.loc[is_sappel | is_itron, 'Anomalie'] += fp2e_anomalies.fillna('').astype(str).str.replace('None', '')

    # Nettoyage de la colonne 'Anomalie'
    df_with_anomalies['Anomalie'] = df_with_anomalies['Anomalie'].str.strip().str.rstrip('/')
    
    anomalies_df = df_with_anomalies[df_with_anomalies['Anomalie'] != ''].copy()
    
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
                "Coordonnées invalides": ['Latitude', 'Longitude'],
                "Diamètre manquant": ['Diametre'],
                
                # Anomalies spécifiques
                "KAMSTRUP: Compteur ≠ Tête": ['Numéro de compteur', 'Numéro de tête'],
                "KAMSTRUP: Compteur ou Tête non numérique": ['Numéro de compteur', 'Numéro de tête'],
                "SAPPEL: Tête ≠ 16 caractères": ['Numéro de tête'],
                "SAPPEL: Compteur format incorrect": ['Numéro de compteur'],
                "ITRON: Tête ≠ 8 caractères": ['Numéro de tête'],
                "ITRON: Compteur doit commencer par \"I\" ou \"D\"": ['Numéro de compteur'],
                "SAPPEL: Incohérence Marque/Compteur (C)": ['Marque', 'Numéro de compteur'],
                "SAPPEL: Incohérence Marque/Compteur (H)": ['Marque', 'Numéro de compteur'],
                "Protocole ≠ LRA pour Traité 903/863": ['Protocole Radio', 'Traité'],
                "Protocole ≠ SGX pour Traité non 903/863": ['Protocole Radio', 'Traité'],

                # Anomalies FP2E
                "Année dans compteur ≠ Année de fabrication": ['Numéro de compteur', 'Année de fabrication'],
                "Lettre de diamètre incohérente": ['Numéro de compteur', 'Diametre'],
                "Règle FP2E non applicable": ['Numéro de compteur', 'Année de fabrication', 'Diametre'],
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
                anomalies_df.to_excel(excel_buffer, index=False, sheet_name='Anomalies', engine='openpyxl')
                excel_buffer.seek(0)
                
                wb = load_workbook(excel_buffer)
                ws = wb.active
                
                red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

                for i, row in enumerate(anomalies_df.iterrows()):
                    anomalies = str(row[1]['Anomalie']).split(' / ')
                    for anomaly in anomalies:
                        # Nettoyer l'anomalie pour la faire correspondre à la clé du dictionnaire
                        anomaly_key = anomaly.strip()
                        if anomaly_key in anomaly_columns_map:
                            columns_to_highlight = anomaly_columns_map[anomaly_key]
                            for col_name in columns_to_highlight:
                                try:
                                    col_index = list(anomalies_df.columns).index(col_name) + 1
                                    cell = ws.cell(row=i + 2, column=col_index)
                                    cell.fill = red_fill
                                except ValueError:
                                    pass
                                
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
