from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl

from datetime import datetime
import time
import random
import dictionary
import re

#Seule variable à changer, mois à scrapper en français
scrapping_month = 'Mars'
actual_day = int(datetime.now().day)
verif_date_event = actual_day-1
jours_depuis_dernier_scrapping = 1 #Mettre 1 si scrapping fait hier, ...
date_dernier_scrapping = actual_day - jours_depuis_dernier_scrapping

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
def rename_prompt_to_midjourney(prompt_initial): #Modification du prompt pour correspondre au titre automatique lors de l'enregistrement d'une image Midjourney
    prompt_midjourney = prompt_initial
    prompt_midjourney = prompt_midjourney.replace(" ", "_")
    prompt_midjourney = prompt_midjourney.replace(",", "")
    prompt_midjourney = prompt_midjourney.replace("'", "")
    prompt_midjourney = prompt_midjourney.replace("(", "")
    prompt_midjourney = prompt_midjourney.replace(")", "")
    prompt_midjourney = prompt_midjourney.replace("ö", "o")
    prompt_midjourney = prompt_midjourney.replace("ø", "o")
    prompt_midjourney = prompt_midjourney.replace("ü", "u")
    prompt_midjourney = prompt_midjourney.replace("é", "e")
    prompt_midjourney = prompt_midjourney.replace("è", "e")
    prompt_midjourney = prompt_midjourney.replace("í", "i")
    prompt_midjourney = "jyllles_" + prompt_midjourney
    one_winner_one_line['Prompt_Midjourney'] = prompt_midjourney
    
    
def winner_eventdate_concordance(winner,date_event): #vérifier si 1 winner n'a pas gagné 2 events le même jour
    winner_eventdate_concordance = f"{winner} - {date_event}"
    if winner_eventdate_concordance in winner_and_date_event:
        print(f"ATTENTION : {winner} a remporté plusieurs épreuves le {date_event} - {url_event}")
    winner_and_date_event.append(winner_eventdate_concordance)
    
def prompt_import_product(prompt_initial): #Je modifie le prompt initial pour qu'il corresponde à la remise en forme opérée par wordpress quand j'importe une image
    prompt_for_import_product = prompt_initial
    prompt_for_import_product = prompt_for_import_product.replace("(", "")
    prompt_for_import_product = prompt_for_import_product.replace(")", "")
    prompt_for_import_product = prompt_for_import_product.replace(" ", "-")
    prompt_for_import_product = prompt_for_import_product.replace(",", "")
    prompt_for_import_product = prompt_for_import_product.replace("'", "")
    prompt_for_import_product = prompt_for_import_product.replace("ö", "o")
    prompt_for_import_product = prompt_for_import_product.replace("Ö", "O")
    prompt_for_import_product = prompt_for_import_product.replace("ø", "o")
    prompt_for_import_product = prompt_for_import_product.replace("ò", "o")
    prompt_for_import_product = prompt_for_import_product.replace("ó", "o")
    prompt_for_import_product = prompt_for_import_product.replace("ü", "u")
    prompt_for_import_product = prompt_for_import_product.replace("ú", "u")
    prompt_for_import_product = prompt_for_import_product.replace("é", "e")
    prompt_for_import_product = prompt_for_import_product.replace("É", "E")
    prompt_for_import_product = prompt_for_import_product.replace("è", "e")
    prompt_for_import_product = prompt_for_import_product.replace("ã", "a")
    prompt_for_import_product = prompt_for_import_product.replace("ä", "a")
    prompt_for_import_product = prompt_for_import_product.replace("á", "a")
    prompt_for_import_product = prompt_for_import_product.replace("å", "a")
    prompt_for_import_product = prompt_for_import_product.replace("ý", "y")
    prompt_for_import_product = prompt_for_import_product.replace("ß", "s")
    prompt_for_import_product = prompt_for_import_product.replace("í", "i")
    prompt_for_import_product = prompt_for_import_product.replace("ï", "i")
    prompt_for_import_product = prompt_for_import_product.replace("---", "-")
    
    prompt_for_import_product = prompt_for_import_product + ".png"
    return prompt_for_import_product

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

def recent_winner_prompt(date_event, date_dernier_scrapping, prompt_initial, midjourney_parameters, EVENT_COUNTER): #Identification gagnant depuis le dernier scrapping
    date_number = re.search(r'\d+', date_event)
    date_number_int = int(date_number.group())
    if date_number_int >= date_dernier_scrapping :
        recents_winners_prompt_list.append(f"{EVENT_COUNTER} - {prompt_initial} {midjourney_parameters}")

#----------------------VARIABLES INITIALES----------------------#

actual_year = str(datetime.now().year)
actual_day = str(datetime.now().day)
url = "https://www.les-sports.info/calendrier-sport-2024-p0-62024.html"
nom_feuille = "Data"
nom_feuille_eng = "Data eng"
date_event = "Month 1st"
midjourney_parameters = ",oil painting style, colorful lighting,--ar 1:1 --c 100 --s 965 --v 5.1"
COMPETITION_COUNTER = 0
EVENT_COUNTER = 0
SEUIL_CORRESPONDANCE = 20

#----------------------LISTES----------------------#
one_winner_one_line_list = [] #Liste contenant toutes les infos pour chaque event (OWOL)
urls_event = [] #Permet d'avoir les url si compet' hommes/femmes et/ou mixte
winner_and_date_event = [] #utilisé pour vérifier si 1 winner n'a pas gagné 2 events le même jour (cf def winner_eventdate_concordance(winner,date_event))
events_ok_list = [] #Liste des events passés par les différents tamis
data_for_wordpress_list = []
competition_of_sport_list = [] #De manière bête et méchante je vais ajouter dans cette liste {competition} of {sport}
competition_of_sport_traduction = [] #J'ai ici la traduction de la liste précédente pour que la traduction soit la plus exacte possible
month_fr_list = []
month_eng_list = []

#Je vérifie avec les 2 éléments ci-dessous si je n'ai pas plusieurs fois le même n° pour des events différents (soucis pour les diapos si c'est le cas)
event_number_list = []
occurence_event_number = {}

#Série de listes avec toutes les erreurs possibles afin de les regrouper et simplifier la lecture en fin de scrapping
no_city_list = []
no_winner_identified_list  = []
no_competition_of_sport_translation_list = []
no_event_translation_list = []
just_men_or_women_list = []

#Identification gagnant depuis le dernier scrapping et création automatique du prompt pour créer les cartes
recents_winners_prompt_list = []


#----------------------VERIFS EXCEL----------------------#
#1 - Est-ce que le fichier Excel existe ?
nom_fichier = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS.xlsx'
try:
    classeur_exist = openpyxl.load_workbook(nom_fichier)
    feuille_exist = classeur_exist.active
except FileNotFoundError:
    classeur_exist = None

#----------------------SPORTS EN ENTREE----------------------#
#2 - Je charge l'Excel puis créé 2 listes contenant les valeurs présentes dans les 2 feuilles de Data
try:
    classeur = openpyxl.load_workbook(nom_fichier)
    feuille_data = classeur[nom_feuille]
    feuille_data_eng = classeur[nom_feuille_eng]

    # Je créé une liste plus spécifique contenant les éléments de la colonne A
    Sports = [cell.value for cell in feuille_data['A'] if cell.value is not None]
    Eng_Sports = [str(cell.value).strip() for cell in feuille_data_eng['A'] if cell.value is not None]

except FileNotFoundError:
    print(f"Le fichier {nom_fichier} n'a pas été trouvé.")
except KeyError:
    print(f"La feuille {nom_feuille} n'a pas été trouvée dans le fichier {nom_fichier}.")

#3 - Si ça a fonctionné sans erreur ci-dessus c'est que l'Excel existe et les feuilles aussi.
#Je peux donc créer des listes spécifiques pour chacune des colonnes présentes dans mes feuilles de Data
#Pourquoi je fais ça ? Les données présentes dans Data vont me permettre d'uniformiser les données en sortie en fonction de ce que je vais trouver dans les données de mon scrapping. Ex : "cyclo-cross" et "Cyclo cross" deviendront "Cyclo-cross"
  
#----------------------COMPETITIONS EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne B
Competitions = [cell.value for cell in feuille_data['B'] if cell.value is not None]
Eng_Competitions = [cell.value for cell in feuille_data_eng['B'] if cell.value is not None]

#----------------------EPREUVES EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne C
Events = [cell.value for cell in feuille_data['C'] if cell.value is not None]
Eng_Events = [cell.value for cell in feuille_data_eng['C'] if cell.value is not None]

#----------------------PAYS EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne D
Countries = [cell.value for cell in feuille_data['D'] if cell.value is not None]
Eng_Countries = [cell.value for cell in feuille_data_eng['D'] if cell.value is not None]

#----------------------VILLES EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne E
Cities_list = [cell.value for cell in feuille_data['E'] if cell.value is not None]
Eng_Cities_list = [cell.value for cell in feuille_data_eng['E'] if cell.value is not None]
    
#----------------------DATES EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne F
Dates = [cell.value for cell in feuille_data['F'] if cell.value is not None]
Eng_Dates = [cell.value for cell in feuille_data_eng['F'] if cell.value is not None]

#----------------------ABREVIATIONS EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne H (I pour l'ensemble)
Abr = [cell.value for cell in feuille_data['H'] if cell.value is not None]
Eng_Abr = [cell.value for cell in feuille_data_eng['H'] if cell.value is not None]
All_Abr = [cell.value for cell in feuille_data_eng['I'] if cell.value is not None]

#----------------------COMPETITIONS A IGNORER EN ENTREE----------------------#
# Créer la liste avec les éléments de la colonne G
Ignore_Competitions = [cell.value for cell in feuille_data['G'] if cell.value is not None]


#--------------MISE EN PLACE DE LA TRADUCTION COMPETITION OF SPORT EN TRADUCTION PLUS EXACTE--------------#
#Je récupère toutes les valeurs présentes dans la colonne J de la feuille Data eng(competition + sport)
for cell in feuille_data_eng['J'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        competition_of_sport_list.append(cell.value)

#Je récupère toutes les valeurs présentes dans la colonne K (competition + sport)
for cell in feuille_data_eng['K'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        competition_of_sport_traduction.append(cell.value)
        
#--------------MISE EN PLACE DE LA TRADUCTION DES MOIS--------------#      
#Je récupère toutes les valeurs présentes dans la colonne M de la feuille Data eng
for cell in feuille_data_eng['M'][1:]:
    # Vérifier si la cellule n'est pas vide et ajouter sa valeur à la liste
    if cell.value is not None:
        month_fr_list.append(cell.value)
        
#Je récupère toutes les valeurs présentes dans la colonne N de la feuille Data eng
for cell in feuille_data_eng['N'][1:]:
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
                        #Recherche dans la première partie du tableau initial (dates)
                        competition_date = event_informations[0].text if len(event_informations) > 0 else None
                        
                        #Recherche dans la deuxième partie du tableau initial (Pays/ville)
                        event_country_and_city = event_informations[1].text if len(event_informations) > 0 else None
                        competition_country = event_informations[1].text if len(event_informations) > 0 else None
                        city_first_chance = event_informations[1].text if len(event_informations) > 0 else None
                        
                        #Recherche dans la troisième partie du tableau initial (sport/épreuve/compétition/ville)
                        sport = event_informations[2].text if len(event_informations) > 0 else None
                        sport_competition = event_informations[2].text if len(event_informations) > 0 else None
                        city_second_chance = event_informations[2].text if len(event_informations) > 0 else None
                        #print(sport_competition)
                        
                        # On va chercher un pays précis
                        if competition_country is not None:
                            competition_country = competition_country.lower()
                            Good_country = next((Country for Country in Countries if Country.lower() in competition_country.lower()), "/")
                            
                            if Good_country in Countries:
                                Good_country_fr = Good_country
                                index_pays = Countries.index(Good_country)
                                Good_country_eng = Eng_Countries[index_pays]
                                competition_country = Good_country_eng

                        # On va chercher une ville précise
                        if city_first_chance is not None and city_second_chance is not None :
                            
                            City_first_chance_match = [City_in_list for City_in_list in Cities_list if City_in_list.lower() in city_first_chance.lower()]
                            City_second_chance_match = [City_in_list for City_in_list in Cities_list if City_in_list.lower() in city_second_chance.lower()]
                            
                            if City_first_chance_match :
                                Good_city = max(City_first_chance_match, key=len)
                                if Good_city in Cities_list:
                                    index_city = Cities_list.index(Good_city)
                                    city = Eng_Cities_list[index_city]
                            elif City_second_chance_match :
                                Good_city = max(City_second_chance_match, key=len)
                                if Good_city in Cities_list:
                                    index_city = Cities_list.index(Good_city)
                                    city = Eng_Cities_list[index_city]
                            else:
                                city = ""
                                no_city_list.append(f"{city_first_chance} ou {city_second_chance}")
                                #print(f"PAS DE VILLE :  {city_first_chance} - {city_second_chance} ?")
                                
                        else:
                            print("je sais pas ou on est :) ")

                        # On va chercher un sport précis
                        if sport is not None:
                            sport = sport.lower()
                            sport_matches = [Sport for Sport in Sports if Sport.lower() in sport.lower()]
                            if sport_matches:
                                Good_sport = max(sport_matches, key=len)
                                
                                if Good_sport in Sports:
                                    index_sport = Sports.index(Good_sport)
                                    Good_sport_eng = Eng_Sports[index_sport]
                                    sport = Good_sport_eng
                            else:
                                sport = f'{sport} A AJOUTER !! - {url_event}'
                                
                        # On va chercher une compétition précise
                        if sport_competition is not None:
                            sport_competition = sport_competition.lower()
                            competition_matches = [Competition for Competition in Competitions if Competition.lower() in sport_competition.lower()]
                            if competition_matches:
                                Good_competition = max(competition_matches, key=len)
                                
                                if Good_competition in Competitions:
                                    index_competition = Competitions.index(Good_competition)
                                    Good_competition_eng = Eng_Competitions[index_competition]
                                    sport_competition = Good_competition_eng
                                
                            else:
                                sport_competition = f'{sport_competition} A AJOUTER !! - {url_event}'

                            #J'ai pour le moment récupéré les informations suivantes : 
                            #     Date de la competition : competition_date,
                            #     Pays de la competition : competition_country,
                            #     Ville de la competition : city,
                            #     Sport : sport,
                            #     Nom de la competition : sport_competition,

#-----------------------Je vais maintenant dans les pages des compétitions
                        sport_event_link = event_informations[2].select_one('a')['href'] if len(event_informations) > 2 and event_informations[2].select_one('a') else None
                        
                        #Intervertir les coms ci-dessous pour faire du scrapping sur une seule page
                        sport_event_link_text = "https://www.les-sports.info/" + sport_event_link if sport_event_link else None
                        #sport_event_link_text = "https://www.les-sports.info/biathlon-coupe-du-monde-femmes-oberhof-resultats-2023-2024-epr131913.html"
                        
                        if "hommes-epm" in sport_event_link_text :
                            url_event_hommes = sport_event_link_text
                            url_event_femmes = sport_event_link_text.replace("hommes-epm", "femmes-epf")
                            url_event_mixte = sport_event_link_text.replace("hommes-epm", "mixte-epi")
                            #https://www.les-sports.info/cyclo-cross-championnats-de-belgique-resultats-2023-2024-hommes-epm130376.html
                            #https://www.les-sports.info/cyclo-cross-championnats-de-belgique-resultats-2023-2024-femmes-epf130376.html
                        else:
                            url_event_hommes = sport_event_link_text
                            url_event_femmes = ""
                            url_event_mixte = ""
                        urls_event = [url_event_hommes, url_event_femmes, url_event_mixte]
                        
                        for url_event in urls_event :
                            if url_event != None and url_event != "" :
                                event_detail = requests.get(url_event)

                                if result.status_code == 200:
                                    all_competition = BeautifulSoup(event_detail.text, "html.parser")
                                
        #---------------------------Je distingue directement si j'ai affaire à des events par équipe (foot, hand) ou bien des sports potentiellement individuels. Je ressort dans tous les cas avec une liste.
            #-----------------------Si le vainqueur est une équipe nationale, la valeur se trouve dans une balise "h3-vainqueur". Je l'ajoute dans ma variable "specific_competition_titles"
                                    #Si cette variable est non nulle, je procède à la BOUCLE 1)
                                    team_competition = None #Remise à zéro
                                    team_competition = all_competition.select(".h3-vainqueur")
                                    
            #-----------------------Si l'event est à destination d'individualités / groupe d'individualités, la valeur se trouve dans une balise "nomargin".
                                    events_ok_list.clear()
                                    all_events_in_page = None #Remise à zéro
                                    all_events_in_page = all_competition.select(".nomargin a")
                                    
                                    
                #-------------------Je vais maintenant devoir faire de nombreux tris afin de faciliter la suite
                
                    #---------------1er tris : je retire tous les events présents dans ma liste "Ignore_Competitions" (qualifications, demi-finale, ...)
                                    #Exemple url pour test : https://www.les-sports.info/saut-a-ski-coupe-du-monde-hommes-innsbruck-berg-isel-resultats-2023-2024-epr132061.html
                                    for event_in_page in all_events_in_page :
                                        if event_in_page is not None : #les 3 tris autont lieu pour tous les éléments qui passent ce if
                                            event_in_page_text = event_in_page.text
                                            if [Ignore_Competition for Ignore_Competition in Ignore_Competitions if Ignore_Competition.lower() in event_in_page_text.lower()]:
                                                #print(f"J'ignore {event_in_page_text} car épreuve dans la liste noire")
                                                pass
                                            else:
                                                events_ok_list.append(event_in_page_text)
                                                #print(f"1er tris passé pour {event_in_page_text}")
                                                
                    #---------------2ème tris : Parmis les éléments qui ont passé le premier tamis et qui sont dans la liste d'events de la compet, je vais maintenant exclure ceux situés avant une class p_14 indiquant un event annulé
                                    #Exemple url pour test : https://www.les-sports.info/saut-a-ski-coupe-du-monde-femmes-zao-resultats-2023-2024-epr132081.html
                                                p_14_elements = all_competition.find_all(class_="p_14")
                                                for p_14_element in p_14_elements:
                                                    if p_14_element is not None and re.search(r"Cette manche a été annulée", p_14_element.text, re.IGNORECASE):
                                                        event_canceled = p_14_element.find_previous_sibling(class_="nomargin")
                                                        event_canceled = event_canceled.text
                                                        if event_canceled == event_in_page_text :
                                                            events_ok_list.remove(event_canceled)
                                                            #print(f"Event annulé : {event_canceled}. Jeter un oeil sur {url_event}")
                                                        else:
                                                            pass #Si plusieurs events annulés sur la même page, je vais passer par ici dès le premier remove et à chaque boucle
                                                    else:
                                                        pass #Pas de manche annulée
                                                    
                    #---------------3ème tris : je retire les events présents dans un toggle (résultats détaillés donc donnée en double)
                                    #Exemple url pour test : https://www.les-sports.info/ski-alpin-coupe-du-monde-hommes-wengen-resultats-2023-2024-epr131946.html
                                                if event_in_page.find_parent("ul", class_="toggle"):
                                                    events_ok_list.remove(event_in_page_text)
                                                    #print(f"J'ai retiré {event_in_page_text} de la liste car présent dans un résultat détaillé - {url_event}")
                                                else:
                                                    pass #Pas d'event dans un toggle "Résultat détaillé"
                                        else:
                                            print(f'Aucun event dans la page : {url_event}')
                                            
                    #---------------LISTE FINALE : tous les events ayant passé les 3 tris sont dans la liste "specific_event_titles"
                                    #print(f'La liste finale prise en compte est : {events_ok_list} - {url_event}')
                                    
                #-------------------Le site de scrapping est en Français et chaque titre d'event contient l'event + sa date. Itération dans la liste puis traduction séparée pour les deux infos
                                    for specific_event_title_index, specific_event_title in enumerate(events_ok_list,start=1):

                    #---------------1ère traduction : Je vais traduire chaque event. BDD des events dans l'Excel (feuille Data, cf #3 en dfébut de code)
                                        event_matches = [Event for Event in Events if Event.lower() in specific_event_title.lower()]
                                        
                                        if event_matches: #Si l'épreuve fait partie de la liste des events en français
                                            Good_event = max(event_matches, key=len) #Je prends la plus longue correspondance s'il y en a plusieurs
                                            sport_event = Good_event
                                            index_event = Events.index(sport_event)
                                            Good_event_eng = Eng_Events[index_event] #ATTENTION il faut 1 valeur dans chaque case sinon tout se décale
                                            if len(Good_event) == 2: #Je passe ici si la seule correspondance concerne le sexe du compétiteur
                                                just_men_or_women_list.append(f'{specific_event_title} - {url_event}"')
                                                #print(f'ATTENTION ! - "{specific_event_title}" devient "{Good_event_eng} - {url_event}"') #Soit l'event ne contient que Hommes/Femmes soit il contient aussi un event non présent dans la traduction. LAISSER LE PRINT ACTIF !! 
                                            sport_event = Good_event_eng
                                        else: #L'épreuve ne fait pas partie de la liste acceptée car traduction non fournie encore
                                            events_ok_list.remove(specific_event_title)
                                            no_event_translation_list.append(f"{specific_event_title}")
                                            #print(f'AJOUTER EPREUVE (COL C) pour : "{specific_event_title}" puis relancer')
                                            
                                            
                    #---------------2ème traduction : Je vais traduire chaque date d'event. BDD des dates dans l'Excel (feuille Data, cf #3 en dfébut de code)
                                        date_matches = [Date for Date in Dates if Date.lower() in specific_event_title.lower()]
                                        
                                        if date_matches:
                                            Good_date = max(date_matches, key=len)
                                            date_event = Good_date
                                            index_date = Dates.index(date_event)
                                            Good_date_eng = Eng_Dates[index_date]
                                            date_event = Good_date_eng
                                        else:
                                            print(f'Ligne {EVENT_COUNTER+1} : Soucis de date pour {sport_event} - {url_event}')

    #J'ai fini les vérifications. J'ai ajouté dans mon dict la clé Epreuve_index et Date épreuve_index.
    # Place à l'identification des vainqueurs et à l'ajout de la clé Gagnant_index
    
    
    #----------------------------------------------------------------------------------------BOUCLE1---------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------Vainqueur = Equipe nationale---------------------------------------------------------------------
                                    #Spécifique pour les compétitions nationales par équipe
                                    #Exemple url pour test = https://www.les-sports.info/handball-championnats-d-europe-hommes-2024-epr130056.html
                                    #ATTENTION : Cet url est mise APRES l'identification de la compétition
                                    for winning_team in team_competition:
                                        if winning_team is not None:
                                            winning_team_info = winning_team.find(class_='nodecort')
                                            winning_team_name = winning_team_info['title'] if winning_team_info is not None else None
                                            
                #-------------------Le site de scrapping est en Français. Je vais traduire chaque pays de FR à ANG
                                            if winning_team_name in Countries: #Si Pays dans la liste, je chope la traduction en anglais
                                                index_country = Countries.index(winning_team_name)
                                                Good_country_eng = Eng_Countries[index_country]
                                                winning_team_name = Good_country_eng
                                                
                                                #Je prépare les variables pour le dictionnaire
                                                sport_event = "" #Pas d'event particulier, l'équipe gagne la compétition en elle-même
                                                winner = winning_team_name
                                                date_event = "" #Avoir la date de la finale ne m'intéresse pas vu que l'équipe gagne une compétition de plusieurs jours
                                                winner_country = "/"
                                                city = ""
                                                EVENT_COUNTER +=1
                                                one_winner_one_line = dictionary.add_to_dictionnary(EVENT_COUNTER,competition_date,competition_country,city,sport,sport_competition,sport_event,date_event,winner,winner_country,url_event)
                                            else:
                                                winning_team_name = f'Pays à ajouter dans la liste : {winning_team_name}'
                                            
                                            
                                            competition_of_sport = f"{sport_competition} of {sport}"
                                            if competition_of_sport in competition_of_sport_list:
                                                j = competition_of_sport_list.index(competition_of_sport)
                                                competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                prompt_initial = f'{winner} wins the {competition_of_sport_traduction_value} in {competition_country}'
                                                winner_eventdate_concordance(winner,date_event)
                                                one_winner_one_line['Prompt'] = prompt_initial
                                                event_number_list.append(EVENT_COUNTER)
                                            else:
                                                prompt_initial = f'{winner} wins the {sport_competition} {actual_year} of {sport} in {competition_country}'
                                                winner_eventdate_concordance(winner,date_event)
                                                one_winner_one_line['Prompt'] = prompt_initial
                                                #print(f"MANQUE TRADUCTION COMPETITION OF SPORT : {sport_competition} of {sport} - {url_event}")
                                                no_competition_of_sport_translation_list.append(f'{sport_competition} of {sport} - {url_event}')
                                                event_number_list.append(EVENT_COUNTER)
                                            
                                            one_winner_one_line['Prompt'] = prompt_initial
                                            
                                            #Modification de la variable pour le prompt Midjourney
                                            rename_prompt_to_midjourney(prompt_initial)
                                            prompt_import_product(prompt_initial)
                                            
                                            #Je rajoute une valeur vide pour la colonne "Commentaire". Si j'ajoute un Com après le scrapping, le nom de l'image Midjourney ne sera pas modifiée
                                            one_winner_one_line['Commentaire'] = "-"
                                            
                                            one_winner_one_line_list.append(one_winner_one_line)
                                            
                                            #Je modifie le prompt initial qui va servir à créer l'url de chaque image dans wordpress
                                            prompt_import_product(prompt_initial)
                                            #Je balance ça dans le dictionnaire et sa fonction "import_wordpress"
                                            prompt_for_import_product = prompt_import_product(prompt_initial)
                                            short_winner = create_short_winner(winner)
                                            data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng)
                                            data_for_wordpress_list.append(data_for_wordpress)
                                            
                                            #Identification gagnant depuis le dernier scrapping et création automatique du prompt pour créer les cartes
                                            recent_winner_prompt(date_event, date_dernier_scrapping, prompt_initial, midjourney_parameters, EVENT_COUNTER)
                                            
                                        else:
                                            print(f"Je n'ai pas de gagnant ! : {url_event} ")
                                            
    #----------------------------------------------------------------------------------------BOUCLE2---------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------Vainqueur = Individu(s)--------------------------------------------------------------------------
                                    #Spécifique pour les compétitions avec individualités ou groupe d'individualités
                                    #Exemple url pour test = https://www.les-sports.info/saut-a-ski-coupe-du-monde-femmes-zao-resultats-2023-2024-epr132081.html
                                    for event_in_page_index, event_in_page in enumerate(all_events_in_page,start=1): #Je dois reprendre les données sur toute la page et non pas repartir de la liste des events
                                        
                                        if event_in_page.text in events_ok_list : #Je vérifie par contre tout de suite pour ne continuer qu'avec les events passés par les tamis
                                            sportsmen_table = event_in_page.find_all_next('table', class_='table-style-2', limit=1)
                                            specific_event_title = event_in_page.text
                                            
                    #-----------------------Obligé pour le moment de re-traduire. A voir comment optimiser plus tard
                                            event_matches = [Event for Event in Events if Event.lower() in specific_event_title.lower()]
                                        
                                            if event_matches: #Si l'épreuve fait partie de la liste des events en français
                                                Good_event = max(event_matches, key=len) #Je prends la plus longue correspondance s'il y en a plusieurs
                                                sport_event = Good_event
                                                index_event = Events.index(sport_event)
                                                Good_event_eng = Eng_Events[index_event] #ATTENTION il faut 1 valeur dans chaque case sinon tout se décale
                                                sport_event = Good_event_eng
                    #-----------------------Optimiser cette partie du code
                    
                   #-----------------------Idem pour la date. A voir comment optimiser plus tard
                                            date_matches = [Date for Date in Dates if Date.lower() in specific_event_title.lower()]
                                            
                                            if date_matches:
                                                Good_date = max(date_matches, key=len)
                                                date_event = Good_date
                                                index_date = Dates.index(date_event)
                                                Good_date_eng = Eng_Dates[index_date]
                                                date_event = Good_date_eng
                                            else:
                                                print(f'Ligne {EVENT_COUNTER+1} : Soucis de date pour {sport_event} - {url_event}')
                    #-----------------------Optimiser cette partie du code
                    
                    
                                            if sportsmen_table :
                                                first_row_infos = sportsmen_table[0]
                                                solo_winner = None #Je remets à zéro au cas où
                                                team_winner = None #Je remets à zéro au cas où
                                                
                                                solo_winner = first_row_infos.find('a', class_='nodecort') #Si j'ai un gagnant seul, class = nodecort
                                                team_winner = first_row_infos.find(class_='tdcol-70') #Si j'ai plusieurs gagnants, class = tdcol-70
                                                
                                                if solo_winner is not None:  
                                                    winner_style = "solo"
                                                    winner = solo_winner.text
                                                    winner = winner.strip() #Je retire tous les espaces avant/après
                                                elif team_winner is not None :
                                                    winner_style = "team"
                                                    winner = team_winner.text
                                                    winner = winner.strip() #Je retire tous les espaces avant/après
                                                else:
                                                    date_number = re.search(r'\d+', date_event)
                                                    date_number_int = int(date_number.group())
                                                    if date_number_int > verif_date_event :
                                                            #La date de l'event est supérieur à la date de vérif. Je ne print rien
                                                            winner = None
                                                    else:
                                                        winner = None
                                                        #La date de l'event est antérieur à la date du scrapping. Vérifier l'url pour voir pourquoi on a pas de gagnant
                                                        no_winner_identified_list.append(f"Pas de gagnant : {url_event}")
                                                        #print(f"ATTENTION : Pas de gagnant identifié - {url_event}")

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
                                                        if winner_country in All_Abr: #Si l'abréviation répond déjà à ISO 3
                                                            winner_country_info = True
                                                            pass
                                                        elif winner_country in Abr: #Si l'abréviation est présente en Français, on choisit sa traduction en ISO 3
                                                            index_Abr = Abr.index(winner_country)
                                                            Abr_eng = Eng_Abr[index_Abr]
                                                            winner_country = Abr_eng
                                                            winner_country_info = True
                                                        else:
                                                            print(f'Ajouter "{winner_country}" dans les abréviations FR et mettre la traduction en Eng - {url_event}') #Cas où le pays n'est pas présent dans Abr NI dans Eng_Abr
                                                            winner_country = "RELANCER LE SCRIPT ET AJOUTER LE PAYS AUX DATAS"
                                                            winner_country_info = False
                                                    else:
                                                        winner = winners_info[0] #Je prends la valeur telle qu'elle et je suis averti par le print()
                                                        winner_country = "?"
                                                        winner_country_info = False
                                                        #print(f"Diapositive_{EVENT_COUNTER+1} - Je n'ai pas la nationalité dans le site. Trouver un moyen à terme (en attendant prompt sans nationalité : {url_event}")

                                                    
                                                    EVENT_COUNTER +=1
                                                    one_winner_one_line = dictionary.add_to_dictionnary(EVENT_COUNTER,competition_date,competition_country,city,sport,sport_competition,sport_event,date_event,winner,winner_country,url_event)
                                                    
                                                    if winner_country_info :
                                                        if city == "": #Le prompt varie en fonction de si j'ai identifié une ville ou non
                                                            
                                                            competition_of_sport = f"{sport_competition} of {sport}"
                                                            if competition_of_sport in competition_of_sport_list:
                                                                j = competition_of_sport_list.index(competition_of_sport)
                                                                competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                prompt_initial = f'{winner} ({winner_country}) wins the {sport_event} {competition_of_sport_traduction_value} in {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                event_number_list.append(EVENT_COUNTER)
                                                            else:
                                                                prompt_initial = f'{winner} ({winner_country}) wins the {sport_event} {sport_competition} of {sport} in {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                print(f"MANQUE TRADUCTION COMPETITION OF SPORT : {sport_competition} of {sport} - {url_event}")
                                                                event_number_list.append(EVENT_COUNTER)
                                                            
                                                            #Modification de la variable pour le prompt Midjourney
                                                            rename_prompt_to_midjourney(prompt_initial)
                                                            #Je rajoute une valeur vide pour la colonne "Commentaire". Si j'ajoute un Com après le scrapping, le nom de l'image Midjourney ne sera pas modifiée
                                                            one_winner_one_line['Commentaire'] = "-"
                                                            
                                                            #Je modifie le prompt initial qui va servir à créer l'url de chaque image dans wordpress
                                                            prompt_import_product(prompt_initial)
                                                            #Je balance ça dans le dictionnaire et sa fonction "import_wordpress"
                                                            prompt_for_import_product = prompt_import_product(prompt_initial)
                                                            short_winner = create_short_winner(winner)
                                                            data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng)
                                                            data_for_wordpress_list.append(data_for_wordpress)
                                                            
                                                        else:
                                                            competition_of_sport = f"{sport_competition} of {sport}"
                                                            if competition_of_sport in competition_of_sport_list:
                                                                j = competition_of_sport_list.index(competition_of_sport)
                                                                competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                prompt_initial = f'{winner} ({winner_country}) wins the {sport_event} {competition_of_sport_traduction_value} in {city}, {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                event_number_list.append(EVENT_COUNTER)
                                                            else:
                                                                prompt_initial = f'{winner} ({winner_country}) wins the {sport_event} {sport_competition} of {sport} in {city}, {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                print(f"MANQUE TRADUCTION COMPETITION OF SPORT : {sport_competition} of {sport} - {url_event}")
                                                                event_number_list.append(EVENT_COUNTER)

                                                            #Modification de la variable pour le prompt Midjourney
                                                            rename_prompt_to_midjourney(prompt_initial)
                                                            #Je rajoute une valeur vide pour la colonne "Commentaire". Si j'ajoute un Com après le scrapping, le nom de l'image Midjourney ne sera pas modifiée
                                                            one_winner_one_line['Commentaire'] = "-"
                                                            
                                                            #Je modifie le prompt initial qui va servir à créer l'url de chaque image dans wordpress
                                                            prompt_import_product(prompt_initial)
                                                            #Je balance ça dans le dictionnaire et sa fonction "import_wordpress"
                                                            prompt_for_import_product = prompt_import_product(prompt_initial)
                                                            short_winner = create_short_winner(winner)
                                                            data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng)
                                                            data_for_wordpress_list.append(data_for_wordpress)
                                                            
                                                            
                                                        one_winner_one_line_list.append(one_winner_one_line) #j'ajoute le dictionnaire à ma liste contenant tous les gagnants et leurs infos annexes
                                                    
                                                    else :
                                                        if city == "": #Le prompt varie en fonction de si j'ai identifié une ville ou non
                                                            competition_of_sport = f"{sport_competition} of {sport}"
                                                            if competition_of_sport in competition_of_sport_list:
                                                                j = competition_of_sport_list.index(competition_of_sport)
                                                                competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                prompt_initial = f'{winner} wins the {sport_event} {competition_of_sport_traduction_value} in {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                event_number_list.append(EVENT_COUNTER)
                                                            else:
                                                                prompt_initial = f'{winner} wins the {sport_event} {sport_competition} of {sport} in {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                print(f"MANQUE TRADUCTION COMPETITION OF SPORT : {sport_competition} of {sport} - {url_event}")
                                                                event_number_list.append(EVENT_COUNTER)

                                                            
                                                            #Modification de la variable pour le prompt Midjourney
                                                            rename_prompt_to_midjourney(prompt_initial)
                                                            #Je rajoute une valeur vide pour la colonne "Commentaire". Si j'ajoute un Com après le scrapping, le nom de l'image Midjourney ne sera pas modifiée
                                                            one_winner_one_line['Commentaire'] = "-"
                                                            
                                                            #Je modifie le prompt initial qui va servir à créer l'url de chaque image dans wordpress
                                                            prompt_import_product(prompt_initial)
                                                            #Je balance ça dans le dictionnaire et sa fonction "import_wordpress"
                                                            prompt_for_import_product = prompt_import_product(prompt_initial)
                                                            short_winner = create_short_winner(winner)
                                                            data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng)
                                                            data_for_wordpress_list.append(data_for_wordpress)
                                                            
                                                        else:
                                                            competition_of_sport = f"{sport_competition} of {sport}"
                                                            if competition_of_sport in competition_of_sport_list:
                                                                j = competition_of_sport_list.index(competition_of_sport)
                                                                competition_of_sport_traduction_value = competition_of_sport_traduction[j]
                                                                prompt_initial = f'{winner} wins the {sport_event} {competition_of_sport_traduction_value} in {city}, {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                event_number_list.append(EVENT_COUNTER)
                                                            else:
                                                                prompt_initial = f'{winner} wins the {sport_event} {sport_competition} of {sport} in {city}, {competition_country} on {date_event}'
                                                                winner_eventdate_concordance(winner,date_event)
                                                                one_winner_one_line['Prompt'] = prompt_initial
                                                                print(f"MANQUE TRADUCTION COMPETITION OF SPORT : {sport_competition} of {sport} - {url_event}")
                                                                event_number_list.append(EVENT_COUNTER)

                                                            
                                                            #Modification de la variable pour le prompt Midjourney
                                                            rename_prompt_to_midjourney(prompt_initial)
                                                            #Je rajoute une valeur vide pour la colonne "Commentaire". Si j'ajoute un Com après le scrapping, le nom de l'image Midjourney ne sera pas modifiée
                                                            one_winner_one_line['Commentaire'] = "-"
                                                            
                                                            #Je modifie le prompt initial qui va servir à créer l'url de chaque image dans wordpress
                                                            prompt_import_product(prompt_initial)
                                                            #Je balance ça dans le dictionnaire et sa fonction "import_wordpress"
                                                            prompt_for_import_product = prompt_import_product(prompt_initial)
                                                            short_winner = create_short_winner(winner)
                                                            data_for_wordpress = dictionary.import_wordpress (EVENT_COUNTER,short_winner,winner,sport,sport_competition,sport_event, prompt_for_import_product, actual_year, scrapping_month,prompt_initial, month_eng)
                                                            data_for_wordpress_list.append(data_for_wordpress)
                                                            
                                                        one_winner_one_line_list.append(one_winner_one_line) #j'ajoute le dictionnaire à ma liste contenant tous les gagnants et leurs infos annexes
                                                    
                                                    
                                                    #Identification gagnant depuis le dernier scrapping et création automatique du prompt pour créer les cartes
                                                    recent_winner_prompt(date_event, date_dernier_scrapping, prompt_initial, midjourney_parameters, EVENT_COUNTER)
                                                    
                                                else:
                                                    pass #Si j'atterris ici j'ai déjà eu une alerte via : print(f"Je n'ai pas de gagnant identifiable : {url_event}")
                                                
                                            else:
                                                date_number = re.search(r'\d+', date_event)
                                                date_number_int = int(date_number.group())
                                                if date_number_int > verif_date_event :
                                                        #L'event n'a pas encore eu lieu. Pas de soucis et je ne print rien pour ne pas me polluer'
                                                        winner = None
                                                else:
                                                    winner = None
                                                    #La date de l'event est antérieur à la date du scrapping. Potentiel soucis
                                                    no_winner_identified_list.append(f"{url_event}")
                                                    #print(f"ATTENTION : Pas de gagnant identifié - {url_event}")

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


#Je vérifie si je n'ai pas plusieurs fois le même numéro pour plusieurs events
    for number in event_number_list:
        if number in occurence_event_number:
            occurence_event_number[number] += 1
        else:
            occurence_event_number[number] = 1

# Parcourir le dictionnaire pour trouver les valeurs en double et les imprimer
    for number, nombre_occurrences in occurence_event_number.items():
        if nombre_occurrences > 1:
            print("La valeur", number, "apparaît ", nombre_occurrences, "fois dans la liste.")
    print(f"{len(event_number_list)} épreuves ajoutées dans l'Excel")

#J'imprime en fin de scrapping toutes les erreures ensembles par catégorie afin de faciliter la lecture
    print()
    print(f'----------------------------------------------------------------------------------------------------')
    print("ERREURS IDENTIFIEES : ")
    print(f'----------------------------------------------------------------------------------------------------')
    print()
    
    if no_city_list :
        print(f'Compétition sans ville ? Ajouter en Col E si identifiée(s) ci-dessous : ')
        for no_city in no_city_list:
            print(f" - {no_city}")
        print(f'-------------------------')
        print()
    
    if no_winner_identified_list :
        print(f"Pas de gagnant identifié alors que l'épreuve est déjà passée : ")
        for no_winner_identified in no_winner_identified_list:
            print(f" - {no_winner_identified}")
        print(f'-------------------------')
        print()
    
    if no_competition_of_sport_translation_list :
        print(f'Manque la traduction de competition of sport. Data eng Col J : ')
        for no_competition_of_sport_translation in no_competition_of_sport_translation_list:
            print(f" - {no_competition_of_sport_translation}")
        print(f'-------------------------')
        print()
    
    if no_event_translation_list :
        print(f"Manque la traduction de l'épreuve. Ajouter en Col C : ")
        for no_event_translation in no_event_translation_list:
            print(f" - {no_event_translation}")
        print(f'-------------------------')
        print()
    
    if just_men_or_women_list :
        print(f"Epreuve traduite en Men ou Women. Ajouter traduction de l'épreuve en colonne C si pas normal : ")
        for just_men_or_women in just_men_or_women_list:
            print(f" - {just_men_or_women}")
        print(f'-------------------------')
        print()

    if recents_winners_prompt_list :
        print("\033[4m" + "De nouveaux gagnants depuis le dernier scrapping du : " + "\033[0m", end="")
        print(f'{date_dernier_scrapping} {scrapping_month}')
        print("\033[4m" + "Voici les prompts pour créer les images sur Midjourney : " + "\033[0m", end="")
        print()
        for recents_winners_prompt in recents_winners_prompt_list:
            print(f"{recents_winners_prompt}")
        print(f'--------------------------------------------------')
        print()

#Création de l'Excel
    # Créer un DataFrame pandas à partir de la liste d'événements
    #df = pd.DataFrame(competition_list)
    df1 = pd.DataFrame(one_winner_one_line_list)
    df2 = pd.DataFrame(data_for_wordpress_list)

    if classeur_exist is not None:
        # Ajouter la nouvelle feuille si le fichier existe déjà
        with pd.ExcelWriter(nom_fichier, engine='openpyxl', mode='a') as writer:
            try:
                writer.book.remove(writer.book[actual_day])
            except KeyError:
                pass
            
        with pd.ExcelWriter(nom_fichier, engine='openpyxl', mode='a') as writer:
            try:
                writer.book.remove(writer.book["Data for WP"])
            except KeyError:
                pass
                    
            df1.to_excel(writer, sheet_name=f"{actual_day}", index=False)
            df2.to_excel(writer, sheet_name="Data for WP", index=False)
            print(f"Scrapping terminé !")
    else:
        # Écrire le DataFrame dans un nouveau fichier Excel si le fichier n'existe pas
        #df.to_excel(f'{actual_month}-{actual_year}.xlsx', sheet_name=f"{actual_day}", index=False)
        print("Fichier Excel créé avec succès.")
else:
    print("Erreur", result.status_code)