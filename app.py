import streamlit as st
import pandas as pd
import io
import csv
import re

# Table de correspondance Diametre -> Lettre pour FP2E
diametre_lettre = {
    15: ['A', 'U', 'Y', 'Z'],
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
    try:
        sample = file.read(2048).decode('utf-8')
        dialect = csv.Sniffer().sniff(sample)
        file.seek(0)
        return dialect.delimiter
    except Exception:
        file.seek(0)
        return ','

def check_data(df):
    df_with_anomalies = df.copy()

    required_columns = ['Protocole Radio', 'Marque', 'Numéro de compteur', 'Numéro de tête', 'Latitude', 'Longitude', 'Année de fabrication', 'Diametre', 'Traité']
    if not all(col in df_with_anomalies.columns for col in required_columns):
        missing = [col for col in required_columns if col not in df_with_anomalies.columns]
        st.error(f"Colonnes manquantes : {', '.join(missing)}")
        st.stop()

    df_with_anomalies['Anomalie'] = ''

    for idx, row in df_with_anomalies.iterrows():
        marque = row['Marque']
        compteur = str(row['Numéro de compteur'])
        tete = str(row['Numéro de tête'])
        annee = str(row['Année de fabrication'])
        diam = row['Diametre']
        radio = row['Protocole Radio']
        traite = str(row['Traité'])

        # Champs vides
        for col in ['Protocole Radio', 'Marque', 'Numéro de compteur', 'Numéro de tête']:
            if pd.isna(row[col]) or str(row[col]).strip() == '':
                df_with_anomalies.at[idx, 'Anomalie'] += f"Colonne '{col}' vide; "

        # Latitude et Longitude
        try:
            lat = float(row['Latitude'])
            lon = float(row['Longitude'])
            if lat == 0:
                df_with_anomalies.at[idx, 'Anomalie'] += "Latitude = 0; "
            if lon == 0:
                df_with_anomalies.at[idx, 'Anomalie'] += "Longitude = 0; "
            if not (-90 <= lat <= 90) or not (-180 <= lon <= 180):
                df_with_anomalies.at[idx, 'Anomalie'] += "Latitude ou Longitude invalide; "
        except:
            df_with_anomalies.at[idx, 'Anomalie'] += "Latitude ou Longitude non numérique; "

        # Format SAPPEL compteur
        if marque in ["SAPPEL (C)", "SAPPEL(C)", "SAPPEL (H)"]:
            if not re.match(r'^[A-Z]{1}\d{2}[A-Z]{2}\d{6}$', compteur):
                df_with_anomalies.at[idx, 'Anomalie'] += "Format compteur SAPPEL invalide; "

        # Numéro de tête longueur et préfixe
        if marque in ["SAPPEL (C)", "SAPPEL(C)", "SAPPEL (H)"]:
            if len(tete) != 16:
                df_with_anomalies.at[idx, 'Anomalie'] += "Numéro de tête != 16 caractères; "
            if marque in ["SAPPEL (C)", "SAPPEL(C)"] and not tete.startswith("C"):
                df_with_anomalies.at[idx, 'Anomalie'] += "Tête ne commence pas par C pour SAPPEL (C); "
            if marque == "SAPPEL (H)" and not tete.startswith("H"):
                df_with_anomalies.at[idx, 'Anomalie'] += "Tête ne commence pas par H pour SAPPEL (H); "

        # FP2E
        if marque in ["SAPPEL (C)", "SAPPEL (H)", "ITRON"] and len(compteur) >= 5:
            if marque == "SAPPEL (C)" and not compteur.startswith("C"):
                df_with_anomalies.at[idx, 'Anomalie'] += "Compteur doit commencer par C pour SAPPEL (C); "
            if marque == "SAPPEL (H)" and not compteur.startswith("H"):
                df_with_anomalies.at[idx, 'Anomalie'] += "Compteur doit commencer par H pour SAPPEL (H); "
            if marque == "ITRON" and compteur[0] not in ["C", "H"]:
                df_with_anomalies.at[idx, 'Anomalie'] += "ITRON doit commencer par C ou H; "
            if not compteur[1:3].isdigit() or compteur[1:3] != annee[-2:]:
                df_with_anomalies.at[idx, 'Anomalie'] += "Année de fabrication non cohérente; "
            try:
                lettre_diam = compteur[4]
                lettres_attendues = diametre_lettre.get(int(diam), [])
                if lettre_diam not in lettres_attendues:
                    df_with_anomalies.at[idx, 'Anomalie'] += f"Lettre '{lettre_diam}' ne correspond pas au diamètre {diam}; "
            except:
                df_with_anomalies.at[idx, 'Anomalie'] += "Diamètre non valide; "

        # ITRON num tête
        if marque == "ITRON":
            if not tete.startswith("I") and not tete.startswith("D"):
                df_with_anomalies.at[idx, 'Anomalie'] += "Tête ITRON doit commencer par I ou D; "
            if not (len(tete) == 8 and tete.isdigit()):
                df_with_anomalies.at[idx, 'Anomalie'] += "Tête ITRON doit avoir 8 chiffres; "

        # KAMSTRUP
        if marque == "KAMSTRUP":
            if compteur != tete:
                df_with_anomalies.at[idx, 'Anomalie'] += "KAMSTRUP : compteur différent de tête; "
            if not compteur.isdigit() or not tete.isdigit():
                df_with_anomalies.at[idx, 'Anomalie'] += "KAMSTRUP : compteur ou tête contient des lettres; "

        # Préfixe compteur impose la marque
        if compteur.startswith("C") and marque not in ["SAPPEL (C)", "SAPPEL(C)"]:
            df_with_anomalies.at[idx, 'Anomalie'] += "Compteur commence par C mais marque incorrecte; "
        if compteur.startswith("H") and marque != "SAPPEL (H)":
            df_with_anomalies.at[idx, 'Anomalie'] += "Compteur commence par H mais marque incorrecte; "

        # Protocole Radio selon Traité
        if traite.startswith("903") or traite.startswith("863"):
            if radio != "LRA":
                df_with_anomalies.at[idx, 'Anomalie'] += "Traité commence par 903/863 mais radio != LRA; "
        else:
            if radio != "SGX":
                df_with_anomalies.at[idx, 'Anomalie'] += "Traité ne commence pas par 903/863 mais radio != SGX; "

    df_with_anomalies['Anomalie'] = df_with_anomalies['Anomalie'].str.strip().str.rstrip(';')
    anomalies_df = df_with_anomalies[df_with_anomalies['Anomalie'] != '']
    return anomalies_df

# --- Interface Streamlit ---
st.title("Contrôle des données de Télérelève")
st.markdown("Veuillez téléverser votre fichier pour lancer les contrôles.")

uploaded_file = st.file_uploader("Choisissez un fichier", type=['csv', 'xlsx'])

if uploaded_file is not None:
    st.success("Fichier chargé avec succès !")
    try:
        ext = uploaded_file.name.split('.')[-1]
        if ext == 'csv':
            delimiter = get_csv_delimiter(uploaded_file)
            df = pd.read_csv(uploaded_file, sep=delimiter)
        elif ext == 'xlsx':
            df = pd.read_excel(uploaded_file)
        else:
            st.error("Format de fichier non pris en charge.")
            st.stop()
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        st.stop()

    st.subheader("Aperçu du fichier")
    st.dataframe(df.head())

    if st.button("Lancer les contrôles"):
        anomalies_df = check_data(df)

        if not anomalies_df.empty:
            st.error("Anomalies détectées !")
            st.dataframe(anomalies_df)

            if ext == 'csv':
                csv_file = anomalies_df.to_csv(index=False, sep=delimiter).encode('utf-8')
                st.download_button("Télécharger les anomalies (CSV)", data=csv_file, file_name="anomalies.csv", mime='text/csv')
            else:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    anomalies_df.to_excel(writer, index=False)
                st.download_button("Télécharger les anomalies (Excel)", data=buffer.getvalue(), file_name="anomalies.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.success("Aucune anomalie détectée !")
