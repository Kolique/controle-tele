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
        st.error(f"Colonnes manquantes : {', '.join(missing)}")
        st.stop()

    df_with_anomalies['Anomalie'] = ''

    for idx, row in df_with_anomalies.iterrows():
        marque = str(row['Marque']) # Convertir en string pour éviter les problèmes avec pd.isna
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
        # 'Numéro de tête' est maintenant géré spécifiquement pour KAMSTRUP plus bas.
        for col in ['Protocole Radio', 'Marque', 'Numéro de compteur']:
            if pd.isna(row[col]) or str(row[col]).strip() == '':
                log_anomaly(f"Colonne '{col}' vide")

        # Règle spécifique pour 'Numéro de tête' : ne pas considérer vide comme anomalie pour KAMSTRUP
        if marque != "KAMSTRUP":
            if pd.isna(row['Numéro de tête']) or str(row['Numéro de tête']).strip() == '':
                log_anomaly("Colonne 'Numéro de tête' vide")


        # Vérification de la Latitude et Longitude
        try:
            lat = float(row['Latitude'])
            lon = float(row['Longitude'])
            if lat == 0:
                log_anomaly("Latitude = 0")
            if lon == 0:
                log_anomaly("Longitude = 0")
            if not (-90 <= lat <= 90) or not (-180 <= lon <= 180):
                log_anomaly("Latitude ou Longitude invalide")
        except ValueError: # Utiliser ValueError pour les erreurs de conversion
            log_anomaly("Latitude ou Longitude non numérique")

        # Règles spécifiques pour SAPPEL (C), SAPPEL (H)
        if marque in ["SAPPEL (C)", "SAPPEL(C)", "SAPPEL (H)"]:
            if not re.match(r'^[A-Z]{1}\d{2}[A-Z]{2}\d{6}$', compteur):
                log_anomaly("Format compteur SAPPEL invalide")
            if len(tete) != 16:
                log_anomaly("Numéro de tête != 16 caractères")

        # Règles pour SAPPEL (C), SAPPEL (H), ITRON
        if marque in ["SAPPEL (C)", "SAPPEL (H)", "ITRON"] and len(compteur) >= 5:
            if marque == "SAPPEL (C)" and not compteur.startswith("C"):
                log_anomaly("Compteur doit commencer par C pour SAPPEL (C)")
            if marque == "SAPPEL (H)" and not compteur.startswith("H"):
                log_anomaly("Compteur doit commencer par H pour SAPPEL (H)")
            if marque == "ITRON" and compteur[0] not in ["I", "D"]:
                log_anomaly("ITRON : Numéro de compteur doit commencer par I ou D")
            
            try:
                # Assurez-vous que 'diam' est un entier avant de l'utiliser comme clé
                if pd.notna(diam) and isinstance(diam, (int, float)):
                    lettre_diam = compteur[4]
                    lettres_attendues = diametre_lettre.get(int(diam), [])
                    if lettre_diam not in lettres_attendues:
                        log_anomaly(f"Lettre '{lettre_diam}' ne correspond pas au diamètre {int(diam)}")
                else:
                    log_anomaly("Diamètre non valide ou manquant")
            except (IndexError, ValueError): # Gérer les erreurs si compteur[4] n'existe pas ou diam n'est pas un nombre
                log_anomaly("Diamètre non valide ou format de compteur incorrect pour la lettre de diamètre")

        # Règles spécifiques pour ITRON (modifiées pour 8 caractères)
        if marque == "ITRON":
            # Si le numéro de tête n'est pas vide, vérifier sa longueur
            if pd.notna(row['Numéro de tête']) and str(row['Numéro de tête']).strip() != '':
                if len(tete) != 8:
                    log_anomaly("ITRON : Numéro de tête doit faire exactement 8 caractères")

        # Nouvelle règle pour KAMSTRUP (logique de vide déjà gérée, juste la comparaison si présent)
        if marque == "KAMSTRUP":
            # Vérifier si 'Numéro de tête' est présent (non vide ou NaN)
            if pd.notna(row['Numéro de tête']) and str(row['Numéro de tête']).strip() != '':
                # Si le numéro de tête est présent, vérifier la correspondance avec le numéro de compteur
                if compteur != tete:
                    log_anomaly("KAMSTRUP : Numéro de compteur différent du Numéro de tête")
                # Vérifier si compteur ou tête contiennent des lettres (si présents et non vides)
                if not compteur.isdigit() or not tete.isdigit():
                    log_anomaly("KAMSTRUP : Numéro de compteur ou Numéro de tête contient des lettres")
            # Si le numéro de tête est manquant, ce n'est pas une anomalie pour KAMSTRUP, cela est géré par la
            # condition 'if marque != "KAMSTRUP":' pour la vérification générale de colonne vide.


        # Vérification de la marque en fonction du début du numéro de compteur
        if compteur.startswith("C") and marque not in ["SAPPEL (C)", "SAPPEL(C)"]:
            log_anomaly("Compteur commence par C mais marque incorrecte")
        if compteur.startswith("H") and marque != "SAPPEL (H)":
            log_anomaly("Compteur commence par H mais marque incorrecte")

        # Vérification du protocole Radio en fonction du champ 'Traité'
        if traite.startswith("903") or traite.startswith("863"):
            if radio != "LRA":
                log_anomaly("Traité commence par 903/863 mais radio != LRA")
        else:
            if radio != "SGX":
                log_anomaly("Traité ne commence pas par 903/863 mais radio != SGX")

    df_with_anomalies['Anomalie'] = df_with_anomalies['Anomalie'].str.strip().str.rstrip(';')
    anomalies_df = df_with_anomalies[df_with_anomalies['Anomalie'] != '']
    return anomalies_df, anomaly_counter

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

            st.subheader("Récapitulatif des anomalies")
            summary_df = pd.DataFrame(list(anomaly_counter.items()), columns=["Type d'anomalie", "Nombre de cas"])
            st.dataframe(summary_df)

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
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    anomalies_df.to_excel(writer, index=False, sheet_name='Anomalies')
                excel_buffer.seek(0)

                st.download_button(
                    label="Télécharger les anomalies en Excel",
                    data=excel_buffer,
                    file_name='anomalies_telerelève.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
        else:
            st.success("Aucune anomalie détectée ! Les données sont conformes.")
