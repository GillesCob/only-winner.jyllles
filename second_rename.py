import os
import pandas as pd
from datetime import datetime
import shutil


actual_day = str(datetime.now().day)
actual_year = str(datetime.now().year)
scrapping_month = "Mars"
Feuille_datas = str(actual_day)

# Chemin vers le fichier Excel
excel_month_file_path = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS 2.xlsx'

# Lire les données Excel en utilisant pandas depuis la feuille du jour
df1 = pd.read_excel(excel_month_file_path, sheet_name=Feuille_datas)

# Dossiers contenant les images initiales, tous les NFT et enfin ceux nouvellement créés
dossier_initial = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/PNG_DIAPO/cartes_avec_images"
dossier_NFT = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/NFT_READY"
dossier_WP = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/NFT_READY/IMPORT WP"

# Parcourir les fichiers dans le dossier d'images
for nom_diapo in os.listdir(dossier_initial):
    chemin_complet_diapo = os.path.join(dossier_initial, nom_diapo)

    # Vérifier si le fichier est un fichier image
    if nom_diapo.endswith((".png")):
        nom_diapo = nom_diapo[:-4]  # Retirer l'extension ".png"
        
        # Vérifier si le nom du fichier correspond à une valeur de la colonne A
        for index, row in df1.iterrows():
            if nom_diapo == str(row['Diapositive name']):
                
                # Obtenir la valeur correspondante de la colonne "Prompt"
                nouveau_nom = str(row['Nom NFT']) + ".png"
                    
                chemin_fichier_renomme = os.path.join(os.path.dirname(dossier_initial),nouveau_nom)
                os.rename(chemin_complet_diapo,chemin_fichier_renomme)
                
                
                dossier_import_WP_JOUR = os.path.join(dossier_WP, (actual_day + " " + scrapping_month))
                os.makedirs(dossier_import_WP_JOUR, exist_ok=True)
                
                shutil.copy(chemin_fichier_renomme, os.path.join(dossier_NFT,nouveau_nom))
                print("Image envoyée dans le dossier NFT")
                
                shutil.move(chemin_fichier_renomme, os.path.join(dossier_import_WP_JOUR,nouveau_nom))
                print("Image envoyée dans le dossier pour l'import WP")
            