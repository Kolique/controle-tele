import streamlit as st
import pandas as pd
import io
import csv

def get_csv_delimiter(file):
    """Détecte le délimiteur d'un fichier CSV."""
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
    Effectue tous les contrôles sur le DataFrame et retourne un DataFrame avec les anomalies.
    """
    df_with_anomalies = df.copy()
    
    # Vérifier la présence des colonnes requises
    required_columns = ['Protocole Radio', 'Marque', 'Numéro de compteur', 'Numéro de tête', 'Latitude', 'Longitude']
    if not all(col in df_with_anomalies.columns for col in required_columns):
        missing_columns = [col for col in required_columns if col not in df_with_anomalies.columns]
        st.error(f"Votre fichier ne contient pas toutes les colonnes requises. Colonnes manquantes : {', '.join(missing_columns)}")
        st.stop()
    
    df_with_anomalies['Anomalie'] = ''

    # 1. Contrôle des cases vides dans la colonne 'Protocole Radio'
    df_with_anomalies.loc[df_with_anomalies['Protocole Radio'].isnull(), 'Anomalie'] += 'Colonne "Protocole Radio" vide; '

    # 2. Contrôle des cases vides dans la colonne 'Marque'
    df_with_anomalies.loc[df_with_anomalies['Marque'].isnull(), 'Anomalie'] += 'Colonne "Marque" vide; '

    # 3. Contrôle des cases vides dans la colonne 'Numéro de compteur'
    df_with_anomalies.loc[df_with_anomalies['Numéro de compteur'].isnull(), 'Anomalie'] += 'Colonne "Numéro de compteur" vide; '
    
    # 4. Contrôle des cases vides dans la colonne 'Numéro de tête'
    df_with_anomalies.loc[df_with_anomalies['Numéro de tête'].isnull(), 'Anomalie'] += 'Colonne "Numéro de tête" vide; '

    # 5. Contrôle des valeurs égales à zéro pour la 'Latitude'
    df_with_anomalies.loc[df_with_anomalies['Latitude'] == 0, 'Anomalie'] += 'Latitude égale à zéro; '
    
    # 6. Contrôle des valeurs égales à zéro pour la 'Longitude'
    df_with_anomalies.loc[df_with_anomalies['Longitude'] == 0, 'Anomalie'] += 'Longitude égale à zéro; '

    # 7. Contrôle de la plage de la Latitude et Longitude
    invalid_lat_lon = (~df_with_anomalies['Latitude'].between(-90, 90, inclusive='both')) | \
                      (~df_with_anomalies['Longitude'].between(-180, 180, inclusive='both'))
    df_with_anomalies.loc[invalid_lat_lon, 'Anomalie'] += "Latitude ou Longitude invalide; "

    # 8. Contrôle de la longueur des caractères pour la marque KAMSTRUP
    kamstrup_condition = (df_with_anomalies['Marque'] == 'KAMSTRUP') & (df_with_anomalies['Numéro de compteur'].astype(str).str.len() != 8)
    df_with_anomalies.loc[kamstrup_condition, 'Anomalie'] += "Marque KAMSTRUP : 'Numéro de compteur' n'a pas 8 caractères; "
    
    # 9. Contrôle de la longueur des caractères pour la marque Sappel(C) ou Sappel(H)
    sappel_condition = (df_with_anomalies['Marque'].isin(['SAPPEL (C)', 'SAPPEL (H)'])) & (df_with_anomalies['Numéro de tête'].astype(str).str.len() != 16)
    df_with_anomalies.loc[sappel_condition, 'Anomalie'] += "Marque SAPPEL (C) ou SAPPEL (H) : 'Numéro de tête' n'a pas 16 caractères; "


    # Nettoyer la colonne d'anomalies
    df_with_anomalies['Anomalie'] = df_with_anomalies['Anomalie'].str.strip().str.rstrip(';')
    
    # Filtrer uniquement les lignes avec des anomalies
    anomalies_df = df_with_anomalies[df_with_anomalies['Anomalie'] != '']
    return anomalies_df

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
        st.error(f"Une erreur est survenue lors de la lecture du fichier : {e}")
        st.stop()

    st.subheader("Aperçu des 5 premières lignes")
    st.dataframe(df.head())

    if st.button("Lancer les contrôles"):
        st.write("Contrôles en cours...")
        anomalies_df = check_data(df)

        if not anomalies_df.empty:
            st.error("Anomalies détectées !")
            st.dataframe(anomalies_df)
            
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
