import streamlit as st
import pandas as pd
import io
import csv
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
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

    # --- DÉBUT DE LA LOGIQUE CORRIGÉE POUR L'ANNÉE DE FABRICATION ---
    # Convertir d'abord la colonne en chaîne de caractères, en remplacant les valeurs manquantes
    # Cela permet de traiter les années comme '8' ou '2008' sans erreur.
    df_with_anomalies['Année de fabrication'] = df_with_anomalies['Année de fabrication'].astype(str).replace('nan', '', regex=False)
    
    # Remplacer les valeurs numériques (y compris celles en float comme '8.0') par un format propre
    # Cette étape est cruciale pour que la transformation en deux chiffres fonctionne bien
    df_with_anomalies['Année de fabrication'] = df_with_anomalies['Année de fabrication'].apply(
        lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() and x != '' else x
    )

    # Convertir l'année en deux chiffres (ex: '2008' -> '08', '8' -> '08')
    df_with_anomalies['Année de fabrication'] = df_with_anomalies['Année de fabrication'].str.slice(-2).str.zfill(2)

    # --- FIN DE LA LOGIQUE CORRIGÉE ---
    
    # Vérification des colonnes requises
    required_columns = ['Protocole Radio', 'Marque', 'Numéro de compteur', 'Numéro de tête', 'Latitude', 'Longitude', 'Année de fabrication', 'Diametre', 'Traité', 'Mode de relève']
    if not all(col in df_with_anomalies.columns for col in required_columns):
        missing = [col for col in required_columns if col not in df_with_anomalies.columns]
        st.error(f"Colonnes requises manquantes : {', '.join(missing)}")
        st.stop()

    df_with_anomalies['Anomalie'] = ''

    # Conversion des colonnes pour les analyses et remplacement des NaN par des chaînes vides
    df_with_anomalies['Numéro de compteur'] = df_with_anomalies['Numéro de compteur'].astype(str).replace('nan', '', regex=False)
    df_with_anomalies['Numéro de tête'] = df_with_anomalies['Numéro de tête'].astype(str).replace('nan', '', regex=False)
    df_with_anomalies['Marque'] = df_with_anomalies['Marque'].astype(str).replace('nan', '', regex=False)
    df_with_anomalies['Protocole Radio'] = df_with_anomalies['Protocole Radio'].astype(str).replace('nan', '', regex=False)
    df_with_anomalies['Traité'] = df_with_anomalies['Traité'].astype(str).replace('nan', '', regex=False)
    df_with_anomalies['Mode de relève'] = df_with_anomalies['Mode de relève'].astype(str).replace('nan', '', regex=False)
    
    # Conversion des colonnes Latitude et Longitude en numérique pour éviter le TypeError
    df_with_anomalies['Latitude'] = pd.to_numeric(df_with_anomalies['Latitude'], errors='coerce')
    df_with_anomalies['Longitude'] = pd.to_numeric(df_with_anomalies['Longitude'], errors='coerce')

    # Marqueurs pour les conditions
    is_kamstrup = df_with_anomalies['Marque'].str.upper() == 'KAMSTRUP'
    is_sappel = df_with_anomalies['Marque'].str.upper().isin(['SAPPEL (C)', 'SAPPEL (H)', 'SAPPEL(C)'])
    is_itron = df_with_anomalies['Marque'].str.upper() == 'ITRON'
    is_mode_manuelle = df_with_anomalies['Mode de relève'].str.upper() == 'MANUELLE'

    annee_fabrication_num = pd.to_numeric(df_with_anomalies['Année de fabrication'], errors='coerce')
    df_with_anomalies['Diametre'] = pd.to_numeric(df_with_anomalies['Diametre'], errors='coerce')

    # ------------------------------------------------------------------
    # ANOMALIES GÉNÉRALES (valeurs manquantes et incohérences de base)
    # ------------------------------------------------------------------
    
    # Règle de l'utilisateur : ne pas considérer comme une anomalie si le protocole radio est manquant
    # ET le mode de relève est "Manuelle".
    condition_protocole_manquant = (df_with_anomalies['Protocole Radio'].isin(['', 'nan'])) & (~is_mode_manuelle)
    df_with_anomalies.loc[condition_protocole_manquant, 'Anomalie'] += 'Protocole Radio manquant / '
    
    # Règle de l'utilisateur : Marque manquante. Cette règle s'applique à tous les cas.
    df_with_anomalies.loc[df_with_anomalies['Marque'].isin(['', 'nan']), 'Anomalie'] += 'Marque manquante / '
    df_with_anomalies.loc[df_with_anomalies['Numéro de compteur'].isin(['', 'nan']), 'Anomalie'] += 'Numéro de compteur manquant / '
    df_with_anomalies.loc[df_with_anomalies['Diametre'].isnull(), 'Anomalie'] += 'Diamètre manquant / '
    df_with_anomalies.loc[annee_fabrication_num.isnull(), 'Anomalie'] += 'Année de fabrication manquante / '
    
    # Règle de l'utilisateur : Numéro de tête manquant est une anomalie SAUF pour Kamstrup ET SAUF si le mode de relève est "Manuelle".
    condition_tete_manquante = (df_with_anomalies['Numéro de tête'].isin(['', 'nan'])) & (~is_kamstrup) & (~is_mode_manuelle)
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
    diametre_kamstrup_anomalie = is_kamstrup & (~df_with_anomalies['Diametre'].between(15, 80))
    df_with_anomalies.loc[diametre_kamstrup_anomalie, 'Anomalie'] += 'KAMSTRUP: Diamètre hors de la plage [15, 80] / '

    # SAPPEL
    sappel_valid = is_sappel & (~df_with_anomalies['Numéro de tête'].isin(['', 'nan']))
    df_with_anomalies.loc[sappel_valid & (df_with_anomalies['Numéro de tête'].str.len() != 16), 'Anomalie'] += 'SAPPEL: Tête ≠ 16 caractères / '
    df_with_anomalies.loc[is_sappel & (~df_with_anomalies['Numéro de compteur'].str.match(r'^[A-Z]{1}\d{2}[A-Z]{2}\d{6}$')), 'Anomalie'] += 'SAPPEL: Compteur format incorrect / '
    df_with_anomalies.loc[(is_sappel) & (df_with_anomalies['Numéro de compteur'].str.startswith('C', na=False)) & (df_with_anomalies['Marque'].str.upper() != 'SAPPEL (C)'), 'Anomalie'] += 'SAPPEL: Incohérence Marque/Compteur (C) / '
    df_with_anomalies.loc[(is_sappel) & (df_with_anomalies['Numéro de compteur'].str.startswith('H', na=False)) & (df_with_anomalies['Marque'].str.upper() != 'SAPPEL (H)'), 'Anomalie'] += 'SAPPEL: Incohérence Marque/Compteur (H) / '
    
    # ITRON
    itron_valid = is_itron & (~df_with_anomalies['Numéro de tête'].isin(['', 'nan']))
    df_with_anomalies.loc[itron_valid & (df_with_anomalies['Numéro de tête'].str.len() != 8), 'Anomalie'] += 'ITRON: Tête ≠ 8 caractères / '
    
    # ------------------------------------------------------------------
    # ANOMALIES MULTI-COLONNES (logique consolidée)
    # ------------------------------------------------------------------

    # Protocole Radio vs Traité
    traite_lra_condition = df_with_anomalies['Traité'].str.startswith(('903', '863'), na=False)
    condition_radio_lra = traite_lra_condition & (df_with_anomalies['Protocole Radio'].str.upper() != 'LRA') & (~is_mode_manuelle)
    df_with_anomalies.loc[condition_radio_lra, 'Anomalie'] += 'Protocole ≠ LRA pour Traité 903/863 / '
    
    condition_radio_sgx = (~traite_lra_condition) & (df_with_anomalies['Protocole Radio'].str.upper() != 'SGX') & (~is_mode_manuelle)
    df_with_anomalies.loc[condition_radio_sgx, 'Anomalie'] += 'Protocole ≠ SGX pour Traité non 903/863 / '

    # Règle de diamètre FP2E (pour SAPPEL et ITRON)
    sappel_itron_fp2e_condition = (is_sappel | is_itron) & \
                                  (df_with_anomalies['Numéro de compteur'].str.len() >= 5) & \
                                  (df_with_anomalies['Diametre'].notnull())

    # Utilisation d'une map plus robuste pour la vérification du diamètre
    fp2e_map = {'A': 15, 'U': 15, 'V': 15, 'B': 20, 'C': 25, 'D': 30, 'E': 40, 'F': 50, 'G': [60, 65], 'H': 80, 'I': 100, 'J': 125, 'K': 150}

    def check_fp2e(row):
        try:
            compteur = row['Numéro de compteur']
            if len(compteur) < 5:
                return False
            
            annee_compteur = compteur[1:3]
            annee_fabrication = row['Année de fabrication']
            
            if pd.isna(annee_fabrication) or annee_fabrication == '' or annee_compteur != annee_fabrication:
                return False
            
            lettre_diam = compteur[4].upper()
            
            expected_diametres = fp2e_map.get(lettre_diam, [])
            
            if not isinstance(expected_diametres, list):
                expected_diametres = [expected_diametres]
            
            return row['Diametre'] in expected_diametres

        except (TypeError, ValueError, IndexError):
            return False

    fp2e_compliant = df_with_anomalies[sappel_itron_fp2e_condition].apply(check_fp2e, axis=1)
    
    # Ajout des messages d'anomalie pour SAPPEL et ITRON
    df_with_anomalies.loc[sappel_itron_fp2e_condition & ~fp2e_compliant & is_sappel, 'Anomalie'] += 'SAPPEL: non conforme FP2E / '
    df_with_anomalies.loc[sappel_itron_fp2e_condition & ~fp2e_compliant & is_itron, 'Anomalie'] += 'ITRON: non conforme FP2E / '
    
    # NOUVELLE LOGIQUE : L'anomalie "ITRON: Compteur doit commencer par "I" ou "D"" est vérifiée
    # UNIQUEMENT si le compteur est un ITRON et qu'il est conforme au format FP2E.
    itron_fp2e_compliant_mask = sappel_itron_fp2e_condition & fp2e_compliant & is_itron
    df_with_anomalies.loc[itron_fp2e_compliant_mask & (~df_with_anomalies['Numéro de compteur'].str.lower().str.startswith(('i', 'd'), na=False)), 'Anomalie'] += 'ITRON: Compteur doit commencer par "I" ou "D" / '

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
        
        # Définir le type de données pour les colonnes pour éviter la notation scientifique
        dtype_mapping = {
            'Numéro de branchement': str,
            'Abonnement': str
        }

        if file_extension == 'csv':
            delimiter = get_csv_delimiter(uploaded_file)
            df = pd.read_csv(uploaded_file, sep=delimiter, dtype=dtype_mapping)
        elif file_extension == 'xlsx':
            df = pd.read_excel(uploaded_file, dtype=dtype_mapping)
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
                
                "KAMSTRUP: Compteur ≠ Tête": ['Numéro de compteur', 'Numéro de tête'],
                "KAMSTRUP: Compteur ou Tête non numérique": ['Numéro de compteur', 'Numéro de tête'],
                "KAMSTRUP: Diamètre hors de la plage [15, 80]": ['Diametre'],
                "SAPPEL: Tête ≠ 16 caractères": ['Numéro de tête'],
                "SAPPEL: Compteur format incorrect": ['Numéro de compteur'],
                "SAPPEL: Incohérence Marque/Compteur (C)": ['Marque', 'Numéro de compteur'],
                "SAPPEL: Incohérence Marque/Compteur (H)": ['Marque', 'Numéro de compteur'],
                "ITRON: Tête ≠ 8 caractères": ['Numéro de tête'],
                "ITRON: Compteur doit commencer par \"I\" ou \"D\"": ['Numéro de compteur'],
                "Protocole ≠ LRA pour Traité 903/863": ['Protocole Radio', 'Traité'],
                "Protocole ≠ SGX pour Traité non 903/863": ['Protocole Radio', 'Traité'],
                "SAPPEL: non conforme FP2E": ['Numéro de compteur', 'Diametre', 'Année de fabrication'],
                "ITRON: non conforme FP2E": ['Numéro de compteur', 'Diametre', 'Année de fabrication'],
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
                
                # Création d'un classeur Excel
                wb = Workbook()
                
                # Suppression de la première feuille par défaut qui est vide et création de la feuille "Récapitulatif"
                default_sheet = wb.active
                wb.remove(default_sheet)
                ws_summary = wb.create_sheet(title="Récapitulatif", index=0)
                
                # Ajout de la nouvelle feuille "Toutes les anomalies"
                ws_all_anomalies = wb.create_sheet(title="Toutes_Anomalies", index=1)
                for r_df_idx, row_data in enumerate(dataframe_to_rows(anomalies_df, index=False, header=True)):
                    ws_all_anomalies.append(row_data)

                # Mise en forme de la feuille "Toutes les anomalies"
                header_font = Font(bold=True)
                red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

                for cell in ws_all_anomalies[1]:
                    cell.font = header_font

                for row_num_all, df_row in enumerate(anomalies_df.iterrows()):
                    anomalies = str(df_row[1]['Anomalie']).split(' / ')
                    for anomaly in anomalies:
                        anomaly_key = anomaly.strip()
                        if anomaly_key in anomaly_columns_map:
                            columns_to_highlight = anomaly_columns_map[anomaly_key]
                            for col_name in columns_to_highlight:
                                try:
                                    col_index = list(anomalies_df.columns).index(col_name) + 1
                                    cell = ws_all_anomalies.cell(row=row_num_all + 2, column=col_index)
                                    cell.fill = red_fill
                                except ValueError:
                                    pass

                # Ajuster la largeur des colonnes dans la feuille "Toutes les anomalies"
                for col in ws_all_anomalies.columns:
                    max_length = 0
                    column = col[0].column # Get the column letter (A, B, C, ...)
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws_all_anomalies.column_dimensions[get_column_letter(column)].width = adjusted_width

                # Mise en forme pour le titre du résumé
                title_font = Font(bold=True, size=16)
                
                ws_summary['A1'] = "Récapitulatif des anomalies"
                ws_summary['A1'].font = title_font
                
                ws_summary.append([]) # Ligne vide pour la séparation
                ws_summary.append(["Type d'anomalie", "Nombre de cas"])
                ws_summary['A3'].font = header_font
                ws_summary['B3'].font = header_font
                
                # Création d'une liste pour stocker les noms de feuilles déjà créées
                created_sheet_names = set(["Toutes_Anomalies"]) # Ajouter le nom de la nouvelle feuille

                # Ajouter un lien vers la nouvelle feuille "Toutes les anomalies"
                ws_summary.cell(row=ws_summary.max_row + 1, column=1, value="Toutes les anomalies").hyperlink = f"#Toutes_Anomalies!A1"
                ws_summary.cell(row=ws_summary.max_row, column=1).font = Font(underline="single", color="0563C1")
                ws_summary.cell(row=ws_summary.max_row, column=2, value=len(anomalies_df)).font = header_font
                ws_summary.cell(row=ws_summary.max_row, column=2).alignment = Alignment(horizontal="right")
                
                for r_idx, (anomaly_type, count) in enumerate(anomaly_counter.items()):
                    # Correction du nettoyage du nom de la feuille et ajout d'une vérification d'unicité
                    sheet_name = re.sub(r'[\\/?*\[\]:()\'"<>|]', '', anomaly_type)
                    sheet_name = sheet_name.replace(' ', '_').replace('.', '').strip()
                    sheet_name = sheet_name[:31] # Tronquer à la longueur max
                    
                    # S'assurer que le nom de la feuille est unique
                    original_sheet_name = sheet_name
                    counter = 1
                    while sheet_name in created_sheet_names:
                        sheet_name = f"{original_sheet_name[:28]}_{counter}"
                        counter += 1
                    created_sheet_names.add(sheet_name)

                    row_num = ws_summary.max_row + 1
                    ws_summary.cell(row=row_num, column=1, value=anomaly_type)
                    ws_summary.cell(row=row_num, column=2, value=count)
                    
                    # Création de la feuille pour cette anomalie
                    ws_anomaly_detail = wb.create_sheet(title=sheet_name)
                    
                    # Écriture des données de l'anomalie dans la feuille dédiée
                    # Filtrer le DataFrame pour ne garder que les lignes contenant cette anomalie
                    filtered_df = anomalies_df[anomalies_df['Anomalie'].str.contains(anomaly_type, regex=False)]
                    
                    # Utilisation de dataframe_to_rows pour écrire les données
                    for r_df_idx, row_data in enumerate(dataframe_to_rows(filtered_df, index=False, header=True)):
                        ws_anomaly_detail.append(row_data)

                    # Mise en forme et en couleur de la feuille détaillée
                    
                    # Mise en surbrillance de la première ligne (en-têtes)
                    for cell in ws_anomaly_detail[1]:
                        cell.font = header_font
                    
                    # Mise en surbrillance des cellules spécifiques
                    for row_num_detail, df_row in enumerate(filtered_df.iterrows()):
                        anomalies = str(df_row[1]['Anomalie']).split(' / ')
                        for anomaly in anomalies:
                            anomaly_key = anomaly.strip()
                            if anomaly_key in anomaly_columns_map:
                                columns_to_highlight = anomaly_columns_map[anomaly_key]
                                for col_name in columns_to_highlight:
                                    try:
                                        # Correction ici : on utilise les colonnes de filtered_df
                                        col_index = list(filtered_df.columns).index(col_name) + 1
                                        cell = ws_anomaly_detail.cell(row=row_num_detail + 2, column=col_index)
                                        cell.fill = red_fill
                                    except ValueError:
                                        pass

                    # Ajuster la largeur des colonnes dans la feuille détaillée
                    for col in ws_anomaly_detail.columns:
                        max_length = 0
                        column = col[0].column # Get the column letter (A, B, C, ...)
                        for cell in col:
                            try: # Necessary to avoid error on non-string cells
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        ws_anomaly_detail.column_dimensions[get_column_letter(column)].width = adjusted_width

                    # Création du lien vers la feuille détaillée sur la page de résumé
                    ws_summary.cell(row=row_num, column=1).hyperlink = f"#{sheet_name}!A1"
                    ws_summary.cell(row=row_num, column=1).font = Font(underline="single", color="0563C1")
                    
                # Ajuster la largeur des colonnes dans le résumé
                for col in ws_summary.columns:
                    max_length = 0
                    column = col[0].column
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws_summary.column_dimensions[get_column_letter(column)].width = adjusted_width
                    
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
