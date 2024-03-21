import os
import pandas as pd
from datetime import datetime
import openpyxl


#ATTENTION : BIEN LIRE LES INFOS SUIVANTES : 
#Va transformer la syntaxe Midjourney en Prompt présents dans l'excel
#Attention à bien mettre à jour le mois concerné !

scrapping_month = "Mars"
actual_day = int(datetime.now().day)
excel_sheet = str(actual_day)


# Chemin vers le fichier Excel
chemin_excel = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS.xlsx'
# Lire les données Excel en utilisant pandas
df = pd.read_excel(chemin_excel, sheet_name=excel_sheet)
# Chemin vers le dossier contenant les images
dossier_images = "/Users/gillescobigo/Downloads"
# Chemin vers le dossier contenant les images renommées
dossier_images_renamed = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/IMAGES_MIDJOURNEY"


classeur_month = openpyxl.load_workbook(chemin_excel)

excel_sheet_informations = classeur_month[excel_sheet]
#Je créé une liste avec les prompts Midjourney
prompt_midjourney_in_excel = [cell.value for cell in excel_sheet_informations['N'] if cell.value is not None]
#Idem avec les noms des cartes
cards_name_in_excel = [cell.value for cell in excel_sheet_informations['M'] if cell.value is not None]

# Parcourir les fichiers dans le dossier d'images
for nom_fichier in os.listdir(dossier_images):
    chemin_fichier = os.path.join(dossier_images, nom_fichier)

    # Vérifier si le fichier est un fichier PNG
    if nom_fichier.lower().endswith(".png"):
        nom_fichier = nom_fichier[:60]
        if nom_fichier in prompt_midjourney_in_excel:
            index_prompt = prompt_midjourney_in_excel.index(nom_fichier)
            nouveau_nom = cards_name_in_excel[index_prompt]
            nouveau_nom = nouveau_nom + ".png"
            os.rename(chemin_fichier, os.path.join(dossier_images_renamed, nouveau_nom))
            
            prompt_midjourney_in_excel.pop(index_prompt)
            cards_name_in_excel.pop(index_prompt)
        else :
            print(f"{nom_fichier} n'est pas dans les prompts Midjourney fournis dans l'excel")
