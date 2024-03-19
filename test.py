from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl

from datetime import datetime
import time
import random
import dictionary
import re

import os

#Seule variable à changer, mois à scrapper en français
scrapping_month = 'Février'
actual_day = int(datetime.now().day)
verif_date_event = actual_day-13
jours_depuis_dernier_scrapping = 1 #Mettre 1 si scrapping fait hier, ...
date_dernier_scrapping = actual_day - jours_depuis_dernier_scrapping
date_event = "March 1st"


#-------------------------------PREREQUIS-------------------------------#
#Copier/coller le dossier "Mois_a_copier" et le nommer avec le mois en cours

#-------------------------------GESTION DU SCRAPPING-------------------------------#
#1er lancement main.py
#Modifier Mois
#Je fais le scrapping jusqu'à la veille du jour actuel : verif_date_event. Si Scrapping sur un mois plus ancien, mettre 32 et non pas {actual_day}

#Prendre en compte l'ensemble des remarques (épreuves à traduire, abréviations pays à ajouter, ...) et les modifier dans l'Excel natif
#Relancer autant de fois qu'il le faudra pour que les erreurs soient validées
#Création d'un excel "mois-2024.xlsx" dans le dossier Excels
#Phase de vérification en fonction des print()
#Vérifier les commentaires présents dans l'Excel (SPORT, COMPETITION, ... à ajouter)
#Relancer jusqu'à ce qu'ils soient réglés ou ignorés
#Ajouter un commentaire en colonne N en fonction des résultats de l'enquête (conséquences dans les points suivants)

#-------------------------------GESTION DES IMAGES-------------------------------#
#Créer le prompt Midjourney dans le Dashboard (suivre les instructions dans l'excel)
#Réaliser les images via Midjourney et les sauvegarder dans le dossier "Download"
#Lancer ren_first_rename.py => Va transformer la syntaxe Midjourney en Prompt présents dans l'excel
#Déplacement auto des images renommées dans le dossier Images du mois scrappé (LES COMPTER)
#GARDE-FOU iMAGES ---------------------------
#Les images ayant un titre trop proche ne sont pas renommés => A traiter à la main
#Si commentaire pour l'event (cf étape à la fin du scrapping) alors image non renommée => A traiter à la main

#-------------------------------GESTION DU PPT-------------------------------#
#Dans le ppt "cartes_sans_images.pptx" créer autant de slides qu'on a d'events dans le l'excel (prendre le n° dans la colonne A)
#Lancer excel_to_ppt.py
#Création automatique du PPT "cartes_avec_images.pptx"
#Ouvrir ce nouveau PPT
#Vérifier la bonne correspondance entre image et prompt
#Vérifier si les épreuves nationales sont masculines ou féminines (refaire l'image si besoin)

#-------------------------------CREATION DES IMAGES FINALES-------------------------------#
#Exporter toutes les slides en .png dans le dossier "PNG_DIAPO"
#Lancer ren_second_rename.py qui va transformer le nom "Diapositive1" en Prompt présents dans l'excel


#------------------------------------------------------------------------------------------------------------------------------------#
#-------------------------------------------------ELEMENTS EN ENTREE-------------------------------------------------#

#----------------------FONCTIONS----------------------#
def rename_prompt_for_midjourney(name_NFT): #Modification du prompt pour correspondre au titre automatique lors de l'enregistrement d'une image Midjourney
    prompt_midjourney = name_NFT
    prompt_midjourney = prompt_midjourney.replace(" ", "_")
    prompt_midjourney = prompt_midjourney.replace(",", "")
    prompt_midjourney = prompt_midjourney.replace("'", "")
    prompt_midjourney = prompt_midjourney.replace("(", "")
    prompt_midjourney = prompt_midjourney.replace(")", "")
    prompt_midjourney = prompt_midjourney.replace("ö", "o")
    prompt_midjourney = prompt_midjourney.replace("ø", "")
    prompt_midjourney = prompt_midjourney.replace("ü", "u")
    prompt_midjourney = prompt_midjourney.replace("é", "e")
    prompt_midjourney = prompt_midjourney.replace("è", "e")
    prompt_midjourney = prompt_midjourney.replace("í", "i")
    prompt_midjourney = prompt_midjourney.replace("ï", "i")
    prompt_midjourney = prompt_midjourney.replace("á", "a")
    prompt_midjourney = "jyllles_" + prompt_midjourney
    new_winners_one_sheet['Prompt_Midjourney'] = prompt_midjourney
    
    
def winner_event_date_concordance(winner,date_event, url_event): #je créé winner-date. S'il est déjà dans la liste, j'append pour le montrer à la fin, sinon je l'ajoute dans la liste
    winner_event_date_concordance = f"{winner} - {date_event}"
    if winner_event_date_concordance in winner_and_date_event_list:
        multiple_winnings_same_day_list.append(f"{winner_event_date_concordance} - {url_event}")
    winner_and_date_event_list.append(winner_event_date_concordance)
    
def prompt_import_product(name_NFT): #Je modifie le prompt initial pour qu'il corresponde à la remise en forme opérée par wordpress quand j'importe une image
    name_NFT = name_NFT + ".png"
    return name_NFT

def card_name_without_accent(name_NFT):
    name_NFT = name_NFT.replace("(", "")
    name_NFT = name_NFT.replace(")", "")
    name_NFT = name_NFT.replace(" ", "-")
    name_NFT = name_NFT.replace(",", "")
    name_NFT = name_NFT.replace("'", "")
    name_NFT = name_NFT.replace("ö", "o")
    name_NFT = name_NFT.replace("Ö", "O")
    name_NFT = name_NFT.replace("ø", "o")
    name_NFT = name_NFT.replace("ò", "o")
    name_NFT = name_NFT.replace("ó", "o")
    name_NFT = name_NFT.replace("ü", "u")
    name_NFT = name_NFT.replace("ú", "u")
    name_NFT = name_NFT.replace("é", "e")
    name_NFT = name_NFT.replace("É", "E")
    name_NFT = name_NFT.replace("è", "e")
    name_NFT = name_NFT.replace("ã", "a")
    name_NFT = name_NFT.replace("ä", "a")
    name_NFT = name_NFT.replace("á", "a")
    name_NFT = name_NFT.replace("å", "a")
    name_NFT = name_NFT.replace("ý", "y")
    name_NFT = name_NFT.replace("ß", "s")
    name_NFT = name_NFT.replace("í", "i")
    name_NFT = name_NFT.replace("ï", "i")
    name_NFT = name_NFT.replace("---", "-")
    name_NFT = name_NFT.replace("--", "-")
    return name_NFT

def create_short_winner(winner): #Diminuer le nombre de caractères afin d'avoir une cohérence visuelle sur la page d'achat des cartes (4 gagnants fait que le titre est trop long et ça décale tout)
    LONGUEUR_TITRE_WORDPRESS = 18
    if len(winner) > LONGUEUR_TITRE_WORDPRESS :
        short_winner = winner.split()
        short_winner = short_winner[:4]
        short_winner = " ".join(short_winner)
        if len(short_winner) > LONGUEUR_TITRE_WORDPRESS :
            short_winner = short_winner[:LONGUEUR_TITRE_WORDPRESS] + "..."
        else:
            pass
    else:
        short_winner = winner
    return short_winner

# def recent_winner_prompt(date_event, date_dernier_scrapping, prompt_initial, midjourney_parameters, EVENT_COUNTER): #Identification gagnant depuis le dernier scrapping
#     if date_event:
#         date_number = re.search(r'\d+', date_event)
#         date_number_int = int(date_number.group())
#         if date_number_int >= date_dernier_scrapping :
#             recents_winners_prompt_list.append(f"{EVENT_COUNTER} - {prompt_initial} {midjourney_parameters}")

#----------------------VARIABLES INITIALES----------------------#

actual_year = str(datetime.now().year)
url = "https://www.les-sports.info/calendrier-sport-2024-p0-62024.html"
midjourney_parameters = ",oil painting style, colorful lighting,--ar 1:1 --c 100 --s 965 --v 5.1"
COMPETITION_COUNTER = 0
EVENT_COUNTER = 0
EVENT_SPECIFIC_COUNTER = 0 #Sert à mettre un numéro qui démarre à 1 pour les nouvelles victoires (nécessaire pour second_rename et l'export PPT)
SEUIL_CORRESPONDANCE = 20

nom_sport_sheet = "SPORT"
nom_competition_sheet = "COMPETITION"
nom_event_sheet = "EVENT"
nom_ignore_event_sheet = "IGNORE EVENT"
nom_country_sheet = "COUNTRY"
nom_city_sheet = "CITY"
nom_date_sheet = "DATE"
nom_abreviation_sheet = "ABREVIATION"
nom_comp_of_sport_sheet = "COMP OF SPORT"
nom_month_sheet = "MONTH"
nom_twitter_sheet = "TWITTER"

dossier_NFT = f"/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/NFT_READY" #Dossier contenant tous les NFT déjà réalisés

    

#----------------------LISTES----------------------#
all_month_winners_list = [] #Liste contenant toutes les infos pour chaque event pour tout le mois
winners_without_nft_list = [] #Contient tous les données concernant les vainqueurs ayant déjà leur carte finalisée

winners_with_nft_list = [] #Contient tous les données concernant les vainqueurs ayant déjà leur carte finalisée

data_for_wordpress_list = [] #Liste de toutes les infos pour l'import Product sur WP

#----------------------LISTES LIEES A DE LA VERIFICATIONS----------------------#
winner_and_date_event_list = [] #j'ajoute dans cette liste les athlètes ayant gagné plusieurs fois le même jour

events_ok_list = [] #Liste des events passés par les différents tamis
all_events_in_page_list = [] #Je mets les events de nomargin et centre

#----------------------LISTES TRADUCTIONS FR EN----------------------#
competition_of_sport_list = [] #De manière bête et méchante je vais ajouter dans cette liste {competition} of {sport}
competition_of_sport_traduction = [] #J'ai ici la traduction de la liste précédente pour que la traduction soit la plus exacte possible

month_fr_list = [] #Traduction du mois afin de faciliter la création des tags produit pour WP
month_eng_list = []

#----------------------LISTES POUR LE PRINT FINAL----------------------#
no_country_list = []
no_city_list = []
no_sport_list = []
no_sport_competition_list = []
no_competition_of_sport_list = []
no_event_list = []
no_date_event_list = []

no_winner_identified_list  = []

no_abr_translation_list = []
just_men_or_women_list = []
recents_winners_prompt_list = []
multiple_winnings_same_day_list = []


#----------------------VERIFS EXCEL----------------------#
#1 - Est-ce que le fichier Excel existe ?
nom_excel_month = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS 2.xlsx'
nom_exel_bdd_datas = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/BDD INITIALE.xlsx'
try:
    classeur_excel_month_exist = openpyxl.load_workbook(nom_excel_month)
    feuille_excel_month_exist = classeur_excel_month_exist.active
except FileNotFoundError:
    classeur_exist = None
    
try:
    classeur_bdd_datas_exist = openpyxl.load_workbook(nom_exel_bdd_datas)
    feuille_bdd_datas_exist = classeur_bdd_datas_exist.active
except FileNotFoundError:
    classeur_exist = None
    
#Je charge tous les prompts pour les NFT déjà réalisés
for nom_nft in os.listdir(dossier_NFT):
    nom_nft = nom_nft[:-4]  # Retirer l'extension ".png"
    winners_with_nft_list.append(nom_nft)
        
#Je charge l'Excel
classeur = openpyxl.load_workbook(nom_excel_month)
classeur_bdd_datas = openpyxl.load_workbook(nom_exel_bdd_datas)

#Ajouts du nom des feuilles Excel suite au changement d'organisation de l'Excel
sport_sheet = classeur_bdd_datas["SPORT"]
competition_sheet = classeur_bdd_datas["COMPETITION"]
event_sheet = classeur_bdd_datas["EVENT"]
ignore_event_sheet = classeur_bdd_datas["IGNORE EVENT"]
country_sheet = classeur_bdd_datas["COUNTRY"]
city_sheet = classeur_bdd_datas["CITY"]
date_sheet = classeur_bdd_datas["DATE"]
abreviation_sheet = classeur_bdd_datas["ABREVIATION"]
comp_of_sport_sheet = classeur_bdd_datas["COMP OF SPORT"]
month_sheet = classeur_bdd_datas["MONTH"]
twitter_sheet = classeur_bdd_datas["TWITTER"]
    


#----------------------SPORTS EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Sports = [cell.value for cell in sport_sheet['A'] if cell.value is not None]
EN_Sports = [cell.value for cell in sport_sheet['B'] if cell.value is not None]

#----------------------COMPETITIONS EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Competition = [cell.value for cell in competition_sheet['A'] if cell.value is not None]
EN_Competition = [cell.value for cell in competition_sheet['B'] if cell.value is not None]

#----------------------EVENT EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Event = [cell.value for cell in event_sheet['A'] if cell.value is not None]
EN_Event = [cell.value for cell in event_sheet['B'] if cell.value is not None]

#----------------------PAYS EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Country = [cell.value for cell in country_sheet['A'] if cell.value is not None]
EN_Country = [cell.value for cell in country_sheet['B'] if cell.value is not None]

#----------------------VILLES EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_City = [cell.value for cell in city_sheet['A'] if cell.value is not None]
EN_City = [cell.value for cell in city_sheet['B'] if cell.value is not None]

#----------------------DATES EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Date = [cell.value for cell in date_sheet['A'] if cell.value is not None]
EN_Date = [cell.value for cell in date_sheet['B'] if cell.value is not None]

#----------------------ABREVIATIONS EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Abreviation = [cell.value for cell in abreviation_sheet['A'] if cell.value is not None]
EN_Abreviation = [cell.value for cell in abreviation_sheet['B'] if cell.value is not None]
ISO3_Abreviation = [cell.value for cell in abreviation_sheet['C'] if cell.value is not None]

#----------------------EVENTS IGNORES EN ENTREE----------------------#
# Je créé deux listes plus spécifiques contenant les éléments de la colonne A (en FR) et B (en ANG)
FR_Ignore_event = [cell.value for cell in ignore_event_sheet['A'] if cell.value is not None]

#----------------------COMPETITION OF SPORT EN ENTREE----------------------#
#Je récupère toutes les valeurs présentes dans la feuille
for cell in comp_of_sport_sheet['A'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        competition_of_sport_list.append(cell.value)

#Je récupère toutes les valeurs présentes dans la colonne K (competition + sport)
for cell in comp_of_sport_sheet['B'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        competition_of_sport_traduction.append(cell.value)
        
#----------------------MOIS EN ENTREE----------------------#
#Je récupère toutes les valeurs présentes dans la colonne A
for cell in month_sheet['A'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        month_fr_list.append(cell.value)
        
#Je récupère toutes les valeurs présentes dans la colonne B
for cell in month_sheet['B'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        month_eng_list.append(cell.value)
        

if scrapping_month in month_fr_list:
    k = month_fr_list.index(scrapping_month)
    month_eng = month_eng_list[k]
else:
    print('SOUCIS DE TRADUCTION MOIS')

#-------------------------------------------------FIN ELEMENTS DE VERIFICATION EN ENTREE-------------------------------------------------#
#----------------------------------------------------------------------------------------------------------------------------------------#

#Appel de mon url contenant les données
result = requests.get(url)

if result.status_code == 200:
    doc = BeautifulSoup(result.text, "html.parser")
    all_months = doc.select('li a.toggle-btn')
    for month in all_months:
        #print(month)
        if month.text == scrapping_month:
            events = month.find_next('div')
            all_events = events.select('tr')
            for event in all_events:
                
                if EVENT_COUNTER<1000 :
                    
                    event_informations = event.select('td')
                    if event_informations :
                        COMPETITION_COUNTER +=1
                        
                        
#---------------------------------------------------------------------------------------------------------------------------------------------------Vérification des infos présentes dans le tableau principal
#-----------------------------------------------------------------------------------On vérifie la bonne traduction des Pays, Villes, Sports, Compétitions

                       
                        #Recherche dans la première partie du tableau initial (dates)
                        competition_date = event_informations[0].text if len(event_informations) > 0 else None
                        
                        #Recherche dans la deuxième partie du tableau initial (Pays/ville)
                        competition_country = event_informations[1].text if len(event_informations) > 0 else None
                        competition_city_first_chance = event_informations[1].text if len(event_informations) > 0 else None
                        
                        #Recherche dans la troisième partie du tableau initial (sport/épreuve/compétition/ville)
                        sport = event_informations[2].text if len(event_informations) > 0 else None
                        sport_competition = event_informations[2].text if len(event_informations) > 0 else None
                        competition_city_second_chance = event_informations[2].text if len(event_informations) > 0 else None


#-----------------------On va chercher un pays précis dans la 2ème colonne du tableau principal
                        if competition_country :
                            Country_match = [Country for Country in FR_Country if Country.lower() in competition_country.lower()]
                            if Country_match:
                                Good_country = max(Country_match, key=len)
                                if Good_country in FR_Country:
                                    Good_country_fr = Good_country
                                    index_pays = FR_Country.index(Good_country)
                                    competition_country = EN_Country[index_pays]
                            else :
                                no_country_list.append(competition_country)
                                competition_country = None
                        else :
                            print("Aucune chance d'arriver ici. Pas de valeur dans la 2ème colonne du tableau principal")

#-----------------------On va chercher une ville précise dans la 2ème et 3ème colonne du tableau principal
                        if competition_city_first_chance and competition_city_second_chance :
                            City_first_chance_match = [City for City in FR_City if City.lower() in competition_city_first_chance.lower()]
                            City_second_chance_match = [City for City in FR_City if City.lower() in competition_city_second_chance.lower()]

                            if City_first_chance_match :
                                Good_city = max(City_first_chance_match, key=len)
                                if Good_city in FR_City:
                                    index_city = FR_City.index(Good_city)
                                    city = EN_City[index_city]
                            elif City_second_chance_match :
                                Good_city = max(City_second_chance_match, key=len)
                                if Good_city in FR_City:
                                    index_city = FR_City.index(Good_city)
                                    city = EN_City[index_city]
                            else:
                                no_city_list.append(f"{competition_city_first_chance} ou {competition_city_second_chance}")
                                city = None
                        else:
                            print("Aucune chance d'arriver ici. Pas de valeur dans la 2ème ou 3ème colonne du tableau principal")


#-----------------------On va chercher un sport précise dans la 3ème colonne du tableau principal
                        if sport :
                            sport_matches = [Sport for Sport in FR_Sports if Sport.lower() in sport.lower()]
                            if sport_matches:
                                Good_sport = max(sport_matches, key=len)
                                if Good_sport in FR_Sports:
                                    index_sport = FR_Sports.index(Good_sport)
                                    Good_sport_eng = EN_Sports[index_sport]
                                    sport = Good_sport_eng.strip()
                            else:
                                no_sport_list.append(sport)
                                sport = None
                        else:
                            print("Aucune chance d'arriver ici. Pas de valeur dans la 3ème colonne du tableau principal")


#-----------------------On va chercher une compétition précise dans la 3ème colonne du tableau principal
                        if sport_competition :
                            sport_competition = sport_competition.lower()
                            competition_matches = [Competition for Competition in FR_Competition if Competition.lower() in sport_competition.lower()]
                            if competition_matches:
                                Good_competition = max(competition_matches, key=len)
                                if Good_competition in FR_Competition:
                                    index_competition = FR_Competition.index(Good_competition)
                                    Good_competition_eng = EN_Competition[index_competition]
                                    sport_competition = Good_competition_eng
                            else:
                                no_sport_competition_list.append(sport_competition)
                                sport_competition = None
                        else:
                            print("Aucune chance d'arriver ici. Pas de valeur dans la 3ème colonne du tableau principal")       



#---------------------------------------------------------------------------------------------------------------------------------------------------Vérification des infos présentes dans les pages d'events
#-----------------------------------------------------------------------------------

                        sport_event_link = event_informations[2].select_one('a')['href'] if len(event_informations) > 2 and event_informations[2].select_one('a') else None
                        sport_event_link_text = "https://www.les-sports.info/" + sport_event_link if sport_event_link else None
                        
                        if "hommes-epm" in sport_event_link_text :
                            url_event_hommes = sport_event_link_text
                            url_event_femmes = sport_event_link_text.replace("hommes-epm", "femmes-epf")
                            url_event_mixte = sport_event_link_text.replace("hommes-epm", "mixte-epi")
                        else:
                            url_event_hommes = sport_event_link_text
                            url_event_femmes = ""
                            url_event_mixte = ""
                        urls_event_list = [url_event_hommes, url_event_femmes, url_event_mixte]
                        
                        for url_event in urls_event_list :
                            if url_event != None and url_event != "" :
                                event_detail = requests.get(url_event)

                                if result.status_code == 200:
                                    all_competition = BeautifulSoup(event_detail.text, "html.parser")


    #----------------------------------------------------------------------------------------BOUCLE1---------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------Vainqueur = Equipe nationale---------------------------------------------------------------------
                                    team_competition = None #Remise à zéro
                                    team_competition = all_competition.select(".h3-vainqueur")
                                    if team_competition :
                                        for winning_team in team_competition:
                                            if winning_team is not None:
                                                EVENT_COUNTER +=1
                                                sport_event = "" #Pas d'event particulier, l'équipe gagne la compétition en elle-même
                                                date_event = "" #Avoir la date de la finale ne m'intéresse pas vu que l'équipe gagne une compétition de plusieurs jours
                                                winner_country = "/" #C'est déjà un pays qui gagne
                                                city = "" #Les compétitions par équipe se passent souvent dans plusieurs villes
                                                
                                                winning_team_info = winning_team.find(class_='nodecort')
                                                winning_team_name = winning_team_info['title'] if winning_team_info is not None else None
                                                
                                                if winning_team_name in FR_Country:
                                                    index_country = FR_Country.index(winning_team_name)
                                                    winner = EN_Country[index_country]
                                                    
                                                    competition_of_sport = f"{sport_competition} of {sport}"
                                                    if competition_of_sport in competition_of_sport_list:
                                                        j = competition_of_sport_list.index(competition_of_sport)
                                                        competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                        
                                                        if competition_country :
                                                            prompt_initial = f'{winner} wins the {competition_of_sport_traduction_value} in {competition_country}'
                                                        else :
                                                            prompt_initial = f'{winner} wins the {competition_of_sport_traduction_value}'
                                                        
                                                        winner_len = str(len(winner))
                                                        name_NFT = f"{winner_len}-{sport}-{sport_competition}-{actual_year}"
                                                        name_NFT = card_name_without_accent(name_NFT)

                                                        #Ci-dessous les éléments spécifiques à intégrer à la feuille contenant tous les évènements du mois
                                                        all_winners_one_sheet = dictionary.add_to_ALL_sheet(competition_date,competition_country,city,sport,sport_competition,sport_event,date_event,winner,winner_country,url_event,prompt_initial,actual_year,name_NFT)
                                                        all_month_winners_list.append(all_winners_one_sheet)
                                                        
                                                        if name_NFT not in winners_with_nft_list : #La carte n'est pas encore créée. j'envoi ces données dans la liste du jour
                                                            #Ci-dessous les éléments spécifiques à intégrer à la feuille contenant les évènements n'ayant pas encore de carte
                                                            EVENT_SPECIFIC_COUNTER +=1
                                                            new_winners_one_sheet = dictionary.add_to_today_sheet(EVENT_SPECIFIC_COUNTER,competition_date,competition_country,city,sport,sport_competition,sport_event,date_event,winner,winner_country,url_event, prompt_initial,actual_year,name_NFT)
                                                            
                                                            rename_prompt_for_midjourney(prompt_initial)
                                                            
                                                            winners_without_nft_list.append(new_winners_one_sheet)
                                                            
                                                            #Ci-dessous les éléments spécifiques à intégrer à la feuille qui permet l'import des produits dans WP
                                                            prompt_for_import_product = prompt_import_product(prompt_initial)
                                                            short_winner = create_short_winner(winner)
                                                            data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng, date_event,name_NFT)
                                                            data_for_wordpress_list.append(data_for_wordpress)
                                                            #J'ai toutes les valeurs pour l'Excel, j'envoi les données du dictionnaire vers la liste qui servira à compléter l'Excel à la date du scrapping
                                                            data_for_wordpress_list.append(data_for_wordpress)
                                                        
                                                    else: #Pas de traduction pour competition of sport donc les infos ne sont pas intégrées à l'excel
                                                        no_competition_of_sport_list.append(f'{sport_competition} of {sport} - {url_event}') #J'ajoute un message final pour rajouter comp of sport dans ma liste Excel
                                                        
                                                else: #Pas de traduction pour le pays gagnant donc les infos ne sont pas intégrées à l'excel
                                                    no_country_list.append(f"{winning_team_name} - {url_event}")

                                            else:
                                                print(f"{winning_team} est None - {url_event }")
                                    
                                    
                                    
        #----------------------------------------------------------------------------------------BOUCLE2---------------------------------------------------------------------------------
        #-------------------------------------------------------------------------------Vainqueur = Individu(s)--------------------------------------------------------------------------                                    
                                    else : #Je n'ai pas de valeur dans h3.vainqueur donc j'arrive ici
                                        events_ok_list.clear()
                                        all_events_in_page = None #Remise à zéro
                                        
                                        all_events_in_page = all_competition.select(".nomargin a")
                                        if all_events_in_page :
                                            pass
                                        else:
                                            all_events_in_page = all_competition.select(".tab_content h2.centre")
                                            

                        #---------------1er tris : je retire tous les events présents dans ma liste "Ignore_Competitions"
                                        for event_in_page in all_events_in_page :
                                            if event_in_page :
                                                event_in_page_text = event_in_page.text
                                                if [Ignore_Competition for Ignore_Competition in FR_Ignore_event if Ignore_Competition.lower() in event_in_page_text.lower()]:
                                                    pass
                                                else:
                                                    events_ok_list.append(event_in_page_text) #Les events non ignorés sont ajoutés à cette liste events_ok_list
                                                    
                        #---------------2ème tris : Parmis les éléments qui ont passé le premier tamis et qui sont dans la liste d'events de la compet, je vais maintenant exclure ceux situés avant une class p_14 indiquant un event annulé
                                                    p_14_elements = all_competition.find_all(class_="p_14") #Je cherche toutes ces classes p_14
                                                    for p_14_element in p_14_elements:
                                                        if p_14_element and re.search(r"Cette manche a été annulée", p_14_element.text, re.IGNORECASE):
                                                            event_canceled = p_14_element.find_previous_sibling(class_="nomargin") #l'event annulé est situé dans la class nomargin située juste avant
                                                            event_canceled = event_canceled.text
                                                            if event_canceled == event_in_page_text : #si l'event avant p_14 est l'event dans lequel on est, je l'enlève de la liste
                                                                events_ok_list.remove(event_canceled)
                                                            else:
                                                                pass 
                                                        else:
                                                            pass #Pas de manche annulée, possible que ce p_14 serve à autre chose
                                                        
                        #---------------3ème tris : je retire les events présents dans un toggle (résultats détaillés donc donnée en double)
                                                    if event_in_page.find_parent("ul", class_="toggle"): #si un event a une class toggle en parent c'est qu'il doit être remove
                                                        events_ok_list.remove(event_in_page_text)
                                                    else:
                                                        pass #Pas d'event dans un toggle "Résultat détaillé"
                                            else:
                                                print(f'Aucun event dans la page : {url_event}')
                                                
        
                                        #J'ai identifié les events à prendre en compte
                                        for event_in_page_index, event_in_page in enumerate(all_events_in_page,start=1): #Je dois reprendre les données sur toute la page et non pas repartir de la liste des events
                                            if event_in_page.text in events_ok_list : #Je vérifie par contre tout de suite pour ne continuer qu'avec les events passés par les tamis
                                                sportsmen_table = event_in_page.find_all_next('table', class_='table-style-2', limit=1)
                                                specific_event_title = event_in_page.text
                                                
                                                
                                                #Je prends le titre d'event et je cherche un event présent dans ma BDD INITIALE
                                                event_matches = [Event for Event in FR_Event if Event.lower() in specific_event_title.lower()]
                                                if event_matches: #Si l'épreuve fait partie de la liste des events en français
                                                    Good_event = max(event_matches, key=len) #Je prends la plus longue correspondance s'il y en a plusieurs
                                                    index_event = FR_Event.index(Good_event)
                                                    sport_event = EN_Event[index_event] #ATTENTION il faut 1 valeur dans chaque case sinon tout se décale
                                                    
                                                    if len(Good_event) == 2: #J'ai créé 1 event ho et 1 event fe (donc si ==2 c'est qu'on a que homme ou femme comme info pour l'event ou bien que l'event n'est pas traduit)
                                                        no_event_list.append(f'{specific_event_title} - {url_event}"')
                                                
                                                else: #L'épreuve ne fait pas partie de la liste acceptée car traduction non fournie encore
                                                    events_ok_list.remove(specific_event_title) #Retirée de la liste des events ok, devra être traduit si on veut son ajout dans l'Excel
                                                    no_event_list.append(f"{specific_event_title} - {url_event}")
                                                    
                        
                                                #Je prends le titre d'event et je cherche une date présente dans ma BDD INITIALE
                                                date_matches = [Date for Date in FR_Date if Date.lower() in specific_event_title.lower()]
                                                if date_matches:
                                                    Good_date = max(date_matches, key=len)
                                                    index_date = FR_Date.index(Good_date)
                                                    date_event = EN_Date[index_date]
                                                else:
                                                    no_date_event_list.append(f"{specific_event_title} - {url_event}")
                                                    
                                                if sportsmen_table :
                                                    first_row_infos = sportsmen_table[0]
                                                    solo_winner = None #Je remets à zéro au cas où
                                                    team_winner = None #Je remets à zéro au cas où
                                                    solo_winner_2 = None #Je remets à zéro au cas où
                                                    solo_winner = first_row_infos.find('a', class_='nodecort') #Si j'ai un gagnant seul, class = nodecort
                                                    team_winner = first_row_infos.find(class_='tdcol-70') #Si j'ai plusieurs gagnants, class = tdcol-70
                                                    
                                                    if solo_winner :  
                                                        winner_style = "solo"
                                                        winner = solo_winner.text
                                                        winner = winner.strip() #Je retire tous les espaces avant/après
                                                    elif team_winner :
                                                        winner_style = "team"
                                                        winner = team_winner.text
                                                        winner = winner.strip() #Je retire tous les espaces avant/après
                                                    else:
                                                        date_number = re.search(r'\d+', date_event)
                                                        date_number_int = int(date_number.group())
                                                        if date_number_int > verif_date_event :
                                                                #La date de l'event est supérieur à la date de vérif. Je ne print rien. Normal de ne pas avoir de gagnant
                                                                winner = None
                                                        else:
                                                            winner = None
                                                            #La date de l'event est antérieur à la date du scrapping. Vérifier l'url pour voir pourquoi on a pas de gagnant
                                                            no_winner_identified_list.append(f"Pas de gagnant : {url_event}")

                                                    #résultat sous la forme "nom (pays)" si "solo" ou "pays (nom1, nom2)" si "team"
                                                    if winner is not None :
                                                        winners_info = winner.split('(')
                                                        if len (winners_info) > 1:
                                                            info_1 = winners_info[0].strip()
                                                            info_2 = winners_info[1].split(')')[0].strip()
                                                            if winner_style == "solo":
                                                                winner = info_1
                                                                winner_country = info_2
                                                            else:
                                                                winner = info_2
                                                                winner_country = info_1
                                                            
                                                            #On va traduire le pays pour qu'il corresponde à la norme ISO 3 et ainsi obtenir un prompt propre et uniformisé
                                                            #url source : https://www.trucsweb.com/tutoriels/internet/iso_3166/
                                                            if winner_country in ISO3_Abreviation: #Si l'abréviation répond déjà à ISO 3
                                                                winner_country_info = True
                                                                pass
                                                            elif winner_country in FR_Abreviation: #Si l'abréviation est présente en Français, on choisit sa traduction en ISO 3
                                                                index_FR_Abr = FR_Abreviation.index(winner_country)
                                                                Abr_eng = EN_Abreviation[index_FR_Abr]
                                                                winner_country = Abr_eng
                                                                winner_country_info = True
                                                            else:
                                                                no_abr_translation_list.append(winner_country)
                                                                winner_country = "RELANCER LE SCRIPT ET AJOUTER LE PAYS AUX DATAS"
                                                                winner_country_info = False
                                                        else:
                                                            winner = winners_info[0] #Je prends la valeur telle qu'elle et je suis averti par le print()
                                                            winner_country = "?"
                                                            winner_country_info = False
                                                            #print(f"Diapositive_{EVENT_COUNTER+1} - Je n'ai pas la nationalité dans le site. Trouver un moyen à terme (en attendant prompt sans nationalité : {url_event}")
                                                        
                                                        EVENT_COUNTER +=1
                                                        sport_competition = sport_competition.strip()
                                                        sport = sport.strip()
                                                        competition_of_sport = f"{sport_competition} of {sport}"
                                                        
                                                        if winner_country_info :
                                                            if city == "": #Le prompt varie en fonction de si j'ai identifié une ville ou non
                                                                if competition_of_sport in competition_of_sport_list:
                                                                    j = competition_of_sport_list.index(competition_of_sport)
                                                                    competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                    prompt_initial = f'{winner} ({winner_country}) wins the {sport_event} {competition_of_sport_traduction_value} in {competition_country} on {date_event}'
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                else:
                                                                    prompt_initial = 'Pas de prompt' #Je dois avoir une formulation de prompt nickel
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                    
                                                                    #ALERTE
                                                                    no_competition_of_sport_list.append(f'{sport_competition} of {sport} - {url_event}')
                                                                
                                                            else:
                                                                if competition_of_sport in competition_of_sport_list:
                                                                    j = competition_of_sport_list.index(competition_of_sport)
                                                                    competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                    prompt_initial = f'{winner} ({winner_country}) wins the {sport_event} {competition_of_sport_traduction_value} in {city}, {competition_country} on {date_event}'
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                else:
                                                                    prompt_initial = 'Pas de prompt' #Je dois avoir une formulation de prompt nickel
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                    
                                                                    #ALERTE
                                                                    no_competition_of_sport_list.append(f'{sport_competition} of {sport} - {url_event}')
                                                                
                                                        else :
                                                            if city == "": #Le prompt varie en fonction de si j'ai identifié une ville ou non
                                                                if competition_of_sport in competition_of_sport_list:
                                                                    j = competition_of_sport_list.index(competition_of_sport)
                                                                    competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                    prompt_initial = f'{winner} wins the {sport_event} {competition_of_sport_traduction_value} in {competition_country} on {date_event}'
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                else:
                                                                    prompt_initial = 'Pas de prompt' #Je dois avoir une formulation de prompt nickel
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                    
                                                                    #ALERTE
                                                                    no_competition_of_sport_list.append(f'{sport_competition} of {sport} - {url_event}')
                                                                
                                                            else:
                                                                if competition_of_sport in competition_of_sport_list:
                                                                    j = competition_of_sport_list.index(competition_of_sport)
                                                                    competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                    prompt_initial = f'{winner} wins the {sport_event} {competition_of_sport_traduction_value} in {city}, {competition_country} on {date_event}'
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                else:
                                                                    prompt_initial = 'Pas de prompt' #Je dois avoir une formulation de prompt nickel
                                                                    winner_event_date_concordance(winner,date_event, url_event)
                                                                    
                                                                    #ALERTE
                                                                    no_competition_of_sport_list.append(f'{sport_competition} of {sport} - {url_event}')

                                                        winner_len = len(winner)
                                                        name_NFT = f"{winner_len}-{sport}-{sport_competition}-{sport_event}-{date_event}-{actual_year}"
                                                        name_NFT = card_name_without_accent(name_NFT)
                                                        
                                                        all_winners_one_sheet = dictionary.add_to_ALL_sheet(competition_date,competition_country,city,sport,sport_competition,sport_event,date_event,winner,winner_country,url_event, prompt_initial,actual_year,name_NFT)
                                                        all_month_winners_list.append(all_winners_one_sheet) #j'ajoute le dictionnaire à ma liste contenant tous les gagnants et leurs infos annexes
                                                        

                                                        if name_NFT not in winners_with_nft_list : #La carte n'est pas encore créée. j'envoi ces données dans la liste du jour
                                                            EVENT_SPECIFIC_COUNTER +=1
                                                            date_number = re.search(r'\d+', date_event)
                                                            date_number_int = int(date_number.group())
                                                            if date_number_int <= actual_day :
                                                                new_winners_one_sheet = dictionary.add_to_today_sheet(EVENT_SPECIFIC_COUNTER,competition_date,competition_country,city,sport,sport_competition,sport_event,date_event,winner,winner_country,url_event, prompt_initial, actual_year,name_NFT)
                                                                rename_prompt_for_midjourney(prompt_initial)
                                                                winners_without_nft_list.append(new_winners_one_sheet)
                                                                
                                                                #Je balance les éléments suivants dans le dictionnaire et sa fonction "import_wordpress"
                                                                prompt_for_import_product = prompt_import_product(name_NFT)
                                                                short_winner = create_short_winner(winner)
                                                                data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng, date_event,name_NFT)
                                                                data_for_wordpress_list.append(data_for_wordpress)
                                                        
                
                                                    else:
                                                        print(f"Je n'ai pas de gagnant identifiable :{sport_competition} - {sport_event} - {url_event}")
                                                    
                                                else:
                                                    date_number = re.search(r'\d+', date_event)
                                                    date_number_int = int(date_number.group())
                                                    if date_number_int > actual_day :
                                                            #L'event n'a pas encore eu lieu. Pas de soucis et je ne print rien pour ne pas me polluer'
                                                            winner = None
                                                    else:
                                                        winner = None
                                                        #La date de l'event est antérieur à la date du scrapping. Potentiel soucis
                                                        no_winner_identified_list.append(f"{url_event}")

                                        pause = random.randrange(1, 3)
                                        time.sleep(pause)   
                                else:
                                    print(f"Pas de retour de l'url. Status code = {result.status_code}")

                                   
    #----------------------------------------------------------------------------------------PLACE A LA CREATION DE L'EXCEL---------------------------------------------------------------------------------
                    #-----------j'ajoute mon dictionnaire dans ma liste d'events
                                #competition_list.append(competition_dict)
                                
                            else:
                                pass #Je passe ici si par exemple je n'ai pas d'url pour les femmes et/ou en mixte
                            
        else:
            pass #On fait quoi si on est pas au actual month ? Bah rien on sort de la boucle et on arrête


#J'imprime en fin de scrapping toutes les erreures ensembles par catégorie afin de faciliter la lecture
    print()
    print(f'----------------------------------------------------------------------------------------------------')
    print("ERREURS IDENTIFIEES : ")
    print(f'----------------------------------------------------------------------------------------------------')
    print() 

    if no_country_list :
        print("\033[4m" + 'Manque les pays suivants dans COUNTRY :' + "\033[0m", end="")
        print()
        for no_country in no_country_list:
            print(f" - {no_country}")
        print(f'-------------------------')
        print()   

    if no_city_list :
        print("\033[4m" + 'Manque les villes suivantes dans CITY :' + "\033[0m", end="")
        print()
        for no_city in no_city_list:
            print(f" - {no_city}")
        print(f'-------------------------')
        print()

    if no_sport_list :
        print("\033[4m" + 'Manque les sports suivants dans SPORT :' + "\033[0m", end="")
        print()
        for no_sport in no_sport_list:
            print(f" - {no_sport}")
        print(f'-------------------------')
        print()
        
    if no_sport_competition_list :
        print("\033[4m" + 'Manque les compétitions suivantes dans COMPETITION :' + "\033[0m", end="")
        print()
        for no_sport_competition in no_sport_competition_list:
            print(f" - {no_sport_competition}")
        print(f'-------------------------')
        print()
        
    if no_competition_of_sport_list :
        print("\033[4m" +'Manque les "compétitions of sport" suivantes dans COMP OD SPORT :' + "\033[0m", end="")
        print()
        for no_competition_of_sport in no_competition_of_sport_list:
            print(f" - {no_competition_of_sport}")
        print(f'-------------------------')
        print()

    if no_event_list :
        print("\033[4m" +'Manque les events suivants dans EVENT (ou bien juste Homme/Femme) :' + "\033[0m", end="")
        print()
        for no_event in no_event_list:
            print(f" - {no_event}")
        print(f'-------------------------')
        print()

    if no_date_event_list :
        print("\033[4m" +"Manque les dates d'event suivantes dans DATE :" + "\033[0m", end="")
        print()
        for no_date_event in no_date_event_list:
            print(f" - {no_date_event}")
        print(f'-------------------------')
        print()
 
   
    
    if no_winner_identified_list :
        print("\033[4m" + "Pas de gagnant identifié alors que l'épreuve est déjà passée : " + "\033[0m", end="")
        print()
        for no_winner_identified in no_winner_identified_list:
            print(f" - {no_winner_identified}")
        print(f'-------------------------')
        print()
    
    if just_men_or_women_list :
        print("\033[4m" + "Epreuve traduite en Men ou Women. Ajouter l'event dans feuille 'EVENT' si pas normal : " + "\033[0m", end="")
        for just_men_or_women in just_men_or_women_list:
            print(f" - {just_men_or_women}")
        print(f'-------------------------')
        print()
        
        
    if no_abr_translation_list:
        print("\033[4m" + "Pas de traduction de l'abréviation d'un pays ? Ajouter dans feuille 'ABREVIATION' : " + "\033[0m", end="")
        print()
        for abr_translation in no_abr_translation_list:
            print(f" - {abr_translation}")
        print(f'-------------------------')
        print()       
        
    if multiple_winnings_same_day_list :
        print("\033[4m" + "Ces athlètes ont remporté plusieurs épreuves le même jour : " + "\033[0m", end="")
        print()
        for winnings_same_day in multiple_winnings_same_day_list:
            print(f" - {winnings_same_day}")
        print(f'-------------------------')
        print()

    if winners_without_nft_list :
        print("\033[4m" + "Voici les prompts pour créer les images sur Midjourney des derniers vainqueurs identifiés " + "\033[0m", end="")
        print()
        for winner_without_nft_list in winners_without_nft_list:
            print()
            print(f"{winner_without_nft_list['Prompt']} {midjourney_parameters}")
        print(f'--------------------------------------------------')
        print()

#Création de l'Excel
    actual_day = str(actual_day)
    # Créer un DataFrame pandas à partir de la liste d'événements
    df1 = pd.DataFrame(all_month_winners_list)
    df2 = pd.DataFrame(winners_without_nft_list)
    df3 = pd.DataFrame(data_for_wordpress_list)

    with pd.ExcelWriter(nom_excel_month, engine='openpyxl', mode='a') as writer:
        try:
            writer.book.remove(writer.book["ALL"])
        except KeyError:
            pass
    with pd.ExcelWriter(nom_excel_month, engine='openpyxl', mode='a') as writer:
        try:
            writer.book.remove(writer.book[actual_day])                  
        except KeyError:
                pass
    with pd.ExcelWriter(nom_excel_month, engine='openpyxl', mode='a') as writer:
        try:
            writer.book.remove(writer.book["Data for WP"])
        except KeyError:
            pass

        df1.to_excel(writer, sheet_name=f"ALL", index=False)
        df2.to_excel(writer, sheet_name=actual_day, index=False)
        df3.to_excel(writer, sheet_name="Data for WP", index=False)

        print(f"Scrapping terminé !")

else:
    print("Erreur", result.status_code)