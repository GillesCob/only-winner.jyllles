import os
import pandas as pd
from datetime import datetime

MOIS_SCRAPPE = "Mars"
Nom_page_prompt_avant_apres = "coucou"
nom_colonne_ancien_prompt = "avant"
nom_colonne_nouveau_prompt = "après"
nouveau_dossier = "renommé"

# Chemin vers le fichier Excel
chemin_excel = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{MOIS_SCRAPPE}/EXCEL/DATAS.xlsx'

# Lire les données Excel en utilisant pandas depuis la feuille "coucou"
df = pd.read_excel(chemin_excel, sheet_name=Nom_page_prompt_avant_apres)

# Dossier contenant les images à renommer
dossier_images = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{MOIS_SCRAPPE}/NFT_READY"

# Dossier de destination pour les images renommées
dossier_images_renamed = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{MOIS_SCRAPPE}/{nouveau_dossier}"

# Parcourir les fichiers dans le dossier d'images
for nom_diapo in os.listdir(dossier_images):
    chemin_fichier = os.path.join(dossier_images, nom_diapo)

    # Vérifier si le fichier est un fichier image
    if nom_diapo.endswith((".png")):
        nom_diapo = nom_diapo[:-4]  # Retirer l'extension ".png"
        # Vérifier si le nom du fichier correspond à une valeur de la colonne A
        for index, row in df.iterrows():
            if nom_diapo == str(row[nom_colonne_ancien_prompt]):
                # Obtenir la valeur correspondante de la colonne B
                nouveau_nom = str(row[nom_colonne_nouveau_prompt]) + ".png"
                # Renommer le fichier image
                os.rename(chemin_fichier, os.path.join(dossier_images_renamed, nouveau_nom))
                print(f"Le fichier {nom_diapo} a été renommé en {nouveau_nom}")