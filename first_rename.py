import os
import pandas as pd
from datetime import datetime

#ATTENTION : BIEN LIRE LES INFOS SUIVANTES : 
#Va transformer la syntaxe Midjourney en Prompt présents dans l'excel
#Si j'ai un commentaire dans la colonne "Commentaire" alors je ne traite pas l'image
#Attention à bien mettre à jour le mois concerné !

actual_month = "Mars"
actual_day = int(datetime.now().day)
excel_sheet = str(actual_day)

titres_images_midjourney = []
CORRESPONDANCE = 50
scrapping_month = actual_month
BOUCLE = 0

# Chemin vers le fichier Excel
chemin_excel = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS 2.xlsx'
# Lire les données Excel en utilisant pandas
df = pd.read_excel(chemin_excel, sheet_name=excel_sheet)
# Chemin vers le dossier contenant les images
dossier_images = "/Users/gillescobigo/Downloads"
# Chemin vers le dossier contenant les images renommées
dossier_images_renamed = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/IMAGES_MIDJOURNEY/VERIF APRES RENOMMAGE"

# Parcourir les fichiers dans le dossier d'images
for nom_fichier in os.listdir(dossier_images):
    chemin_fichier = os.path.join(dossier_images, nom_fichier)

    # Vérifier si le fichier est un fichier PNG
    if nom_fichier.lower().endswith(".png"):
        if not any(titre[:CORRESPONDANCE] == nom_fichier[:CORRESPONDANCE] for titre in titres_images_midjourney): #Je vérifie si 2 fichiers ont le même nom, possible si compet' très semblable et même gagnant
            titres_images_midjourney.append(nom_fichier)
        
            # Lire le titre de l'image
            titre_image = nom_fichier[:-4]  # Retirer l'extension ".png"
            
            # Parcourir les lignes du DataFrame
            for index, row in df.iterrows():
                # Comparer le titre de l'image avec les valeurs de la colonne M
                if titre_image[:CORRESPONDANCE] in str(row['Prompt_Midjourney']):
                    # Modifier le nom du fichier avec la valeur de la colonne L
                    nouveau_nom = str(row['Nom NFT']) + ".png"
                    # Renommer le fichier
                    os.rename(chemin_fichier, os.path.join(dossier_images_renamed, nouveau_nom))
                    #print(f"Le fichier {nom_fichier} a été renommé en {nouveau_nom}")
                    BOUCLE +=1
                    print(f"Ok pour {nouveau_nom}")
                    break  # Arrêter la recherche après la première correspondance trouvée

        else :
            print(f"{nom_fichier} en DOUBLON")
            
#Si l'image est déjà présente dans le dossier "IMAGES_MIDJOURNEY" elle sera écrasée tout simplement