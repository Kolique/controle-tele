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
    60: ['G'], # Valeur G = 60 et G = 65, comme indiqué précédemment
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
    Vérifie les données du DataFrame pour détecter les anomalies.
    Retourne un DataFrame avec les lignes contenant des anomalies et un dictionnaire de comptage des anomalies.
    """
    df_with_anomalies = df.copy()
    anomaly_counter = {}

    # Vérification des colonnes requises
    required_columns = ['Protocole Radio', 'Marque', 'Numéro de compteur', 'Numéro de tête', 'Latitude', 'Longitude', 'Année de fabrication', 'Diametre', 'Traité']
    if not all(col in df_with_anomalies.columns for col in required_columns):
        missing = [col for col in required_columns if col not in df_with_anomalies.columns]
        st.error(f"Colonnes requises manquantes : {', '.join(missing)}")
        st.stop()

    df_with_anomalies['Anomalie'] = ''

    for idx, row in df_with_anomalies.iterrows():
        marque = str(row['Marque'])
        compteur = str(row['Numéro de compteur'])
        tete = str(row['Numéro de tête'])
        annee = str(row['Année de fabrication'])
        diam = row['Diametre']
        radio = str(row['Protocole Radio'])
        traite = str(row['Traité'])

        def log_anomaly(label):
            """Ajoute une anomalie à la ligne actuelle et met à jour le compteur."""
            df_with_anomalies.at[idx, 'Anomalie'] += label + '; '
            anomaly_counter[label] = anomaly_counter.get(label, 0) + 1

        # Vérification des colonnes vides pour les champs essentiels
        for col in ['Protocole Radio', 'Marque', 'Numéro de compteur']:
            if pd.isna(row[col]) or str(row[col]).strip() == '':
                log_anomaly(f"Champ '{col}' manquant")

        # Règle spécifique pour 'Numéro de tête' : ne pas considérer vide comme anomalie pour KAMSTRUP
        if marque != "KAMSTRUP":
            if pd.isna(row['Numéro de tête']) or str(row['Numéro de tête']).strip() == '':
                log_anomaly("Champ 'Numéro de tête' manquant")

        # Vérification de la Latitude et Longitude
        try:
            lat = float(row['Latitude'])
            lon = float(row['Longitude'])
            if lat == 0:
                log_anomaly("Latitude à zéro")
            if lon == 0:
                log_anomaly("Longitude à zéro")
            if not (-90 <= lat <= 90) or not (-180 <= lon <= 180):
                log_anomaly("Coordonnées GPS invalides")
        except ValueError:
            log_anomaly("Coordonnées GPS non numériques")

        # Règles spécifiques pour SAPPEL (C), SAPPEL (H)
        if marque in ["SAPPEL (C)", "SAPPEL(C)", "SAPPEL (H)"]:
            if not re.match(r'^[A-Z]{1}\d{2}[A-Z]{2}\d{6}$', compteur):
                log_anomaly("Format du numéro de compteur SAPPEL incorrect")
            if len(tete) != 16:
                log_anomaly("Longueur du numéro de tête SAPPEL incorrecte (doit être 16 caractères)")

        # Règles pour SAPPEL (C), SAPPEL (H), ITRON
        if marque in ["SAPPEL (C)", "SAPPEL (H)", "ITRON"] and len(compteur) >= 5:
            if marque == "SAPPEL (C)" and not compteur.startswith("C"):
                log_anomaly("Numéro de compteur SAPPEL (C) ne commence pas par 'C'")
            if marque == "SAPPEL (H)" and not compteur.startswith("H"):
                log_anomaly("Numéro de compteur SAPPEL (H) ne commence pas par 'H'")
            if marque == "ITRON" and compteur[0] not in ["I", "D"]:
                log_anomaly("Numéro de compteur ITRON doit commencer par 'I' ou 'D'")
            
            try:
                if pd.notna(diam) and isinstance(diam, (int, float)):
                    lettre_diam = compteur[4]
                    lettres_attendues = diametre_lettre.get(int(diam), [])
                    if lettre_diam not in lettres_attendues:
                        log_anomaly(f"Lettre de diamètre dans le numéro de compteur incohérente")
                else:
                    log_anomaly("Diamètre non valide ou manquant")
            except (IndexError, ValueError):
                log_anomaly("Diamètre non valide ou format de compteur incorrect pour la lettre de diamètre")

        # Règles spécifiques pour ITRON (modifiées pour 8 caractères)
        if marque == "ITRON":
            if pd.notna(row['Numéro de tête']) and str(row['Numéro de tête']).strip() != '':
                if len(tete) != 8:
                    log_anomaly("Longueur du numéro de tête ITRON incorrecte (doit être 8 caractères)")

        # Nouvelle règle pour KAMSTRUP (logique de vide déjà gérée, juste la comparaison si présent)
        if marque == "KAMSTRUP":
            if pd.notna(row['Numéro de tête']) and str(row['Numéro de tête']).strip() != '':
                if compteur != tete:
                    log_anomaly("Numéro de compteur KAMSTRUP différent du Numéro de tête")
                if not compteur.isdigit() or not tete.isdigit():
                    log_anomaly("Numéro de compteur ou de tête KAMSTRUP contient des lettres")

        # Vérification de la marque en fonction du début du numéro de compteur
        if compteur.startswith("C") and marque not in ["SAPPEL (C)", "SAPPEL(C)"]:
            log_anomaly("Marque incohérente avec le numéro de compteur (commence par 'C')")
        if compteur.startswith("H") and marque != "SAPPEL (H)":
            log_anomaly("Marque incohérente avec le numéro de compteur (commence par 'H')")

        # Vérification du protocole Radio en fonction du champ 'Traité'
        if traite.startswith("903") or traite.startswith("863"):
            if radio != "LRA":
                log_anomaly("Protocole Radio incohérent avec le champ 'Traité' (doit être LRA)")
        else:
            if radio != "SGX":
                log_anomaly("Protocole Radio incohérent avec le champ 'Traité' (doit être SGX)")

    df_with_anomalies['Anomalie'] = df_with_anomalies['Anomalie'].str.strip().str.rstrip(';')
    anomalies_df = df_with_anomalies[df_with_anomalies['Anomalie'] != '']
    return anomalies_df, anomaly_counter

def afficher_resume_anomalies(anomaly_counter):
    """
    Affiche un résumé des anomalies.
    """
    if anomaly_counter:
        summary_df = pd.DataFrame(list(anomaly_counter.items()), columns=["Type d'anomalie", "Nombre de cas"])
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
                "Champ 'Protocole Radio' manquant": ['Protocole Radio'],
                "Champ 'Marque' manquant": ['Marque'],
                "Champ 'Numéro de compteur' manquant": ['Numéro de compteur'],
                "Champ 'Numéro de tête' manquant": ['Numéro de tête'],
                "Latitude à zéro": ['Latitude'],
                "Longitude à zéro": ['Longitude'],
                "Coordonnées GPS invalides": ['Latitude', 'Longitude'],
                "Coordonnées GPS non numériques": ['Latitude', 'Longitude'],
                "Format du numéro de compteur SAPPEL incorrect": ['Numéro de compteur'],
                "Longueur du numéro de tête SAPPEL incorrecte (doit être 16 caractères)": ['Numéro de tête'],
                "Numéro de compteur SAPPEL (C) ne commence pas par 'C'": ['Numéro de compteur'],
                "Numéro de compteur SAPPEL (H) ne commence pas par 'H'": ['Numéro de compteur'],
                "Numéro de compteur ITRON doit commencer par 'I' ou 'D'": ['Numéro de compteur'],
                "Lettre de diamètre dans le numéro de compteur incohérente": ['Numéro de compteur', 'Diametre'],
                "Diamètre non valide ou manquant": ['Diametre'],
                "Diamètre non valide ou format de compteur incorrect pour la lettre de diamètre": ['Numéro de compteur', 'Diametre'],
                "Longueur du numéro de tête ITRON incorrecte (doit être 8 caractères)": ['Numéro de tête'],
                "Numéro de compteur KAMSTRUP différent du Numéro de tête": ['Numéro de compteur', 'Numéro de tête'],
                "Numéro de compteur ou de tête KAMSTRUP contient des lettres": ['Numéro de compteur', 'Numéro de tête'],
                "Marque incohérente avec le numéro de compteur (commence par 'C')": ['Marque', 'Numéro de compteur'],
                "Marque incohérente avec le numéro de compteur (commence par 'H')": ['Marque', 'Numéro de compteur'],
                "Protocole Radio incohérent avec le champ 'Traité' (doit être LRA)": ['Protocole Radio', 'Traité'],
                "Protocole Radio incohérent avec le champ 'Traité' (doit être SGX)": ['Protocole Radio', 'Traité']
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
                    anomalies = str(row[1]['Anomalie']).split('; ')
                    for anomaly in anomalies:
                        if anomaly in anomaly_columns_map:
                            columns_to_highlight = anomaly_columns_map[anomaly]
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
