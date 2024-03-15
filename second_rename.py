import os
import pandas as pd
from datetime import datetime
import shutil


actual_day = str(datetime.now().day)
actual_year = str(datetime.now().year)
scrapping_month = "Mars"
Feuille_datas = str(actual_day)
too_long = 1

# Chemin vers le fichier Excel
chemin_excel = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS.xlsx'

# Lire les données Excel en utilisant pandas depuis la feuille du jour
df = pd.read_excel(chemin_excel, sheet_name=Feuille_datas)

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
        for index, row in df.iterrows():
            if nom_diapo == str(row['Nom image']):
                
                # Obtenir la valeur correspondante de la colonne "Prompt"
                nouveau_nom = str(row['Prompt']) + ".png"
                
                if len(nouveau_nom) > 204: #Obligé de mettre une limite car WP en a une
                    ancien_nom = nouveau_nom
                    nouveau_nom = f"{scrapping_month}-{actual_year}-{too_long}.png"
                    print(f"ATTENTION = {ancien_nom}! Penser à changer l'url dans l'import Wordpress")
                    too_long +=1
                    
                    
                chemin_fichier_renomme = os.path.join(os.path.dirname(dossier_initial),nouveau_nom)
                os.rename(chemin_complet_diapo,chemin_fichier_renomme)
                
                dossier_import_WP_JOUR = os.path.join(dossier_WP, (actual_day + " " + scrapping_month))
                os.makedirs(dossier_import_WP_JOUR, exist_ok=True)
                
                shutil.copy(chemin_fichier_renomme, os.path.join(dossier_NFT,nouveau_nom))
                shutil.move(chemin_fichier_renomme, os.path.join(dossier_import_WP_JOUR,nouveau_nom))
            