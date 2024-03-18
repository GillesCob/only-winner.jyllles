from bs4 import BeautifulSoup
import requests

url = "https://www.les-sports.info/tennis-tiers-2-abu-dhabi-500-2024-resultats-eprd133392.html"



#Appel de mon url contenant les données
result = requests.get(url)

if result.status_code == 200:
    all_competition = BeautifulSoup(result.text, "html.parser")
    
    all_events_in_page = all_competition.select(".tab_content h2.centre")
    for event_in_page_index, event_in_page in enumerate(all_events_in_page,start=1): #Je dois reprendre les données sur toute la page et non pas repartir de la liste des events

        sportsmen_table = event_in_page.find_all_next('table', class_='table-style-2', limit=1)
        specific_event_title = event_in_page.text
        
        if sportsmen_table :
            first_row_infos = sportsmen_table[0]
            solo_winner = first_row_infos.find('a', class_='nodecort') #Si j'ai un gagnant seul, class = nodecort
            winner = solo_winner.text
            winner = winner.strip() #Je retire tous les espaces avant/après
            print(winner)
        else :
            print("Pas de tableau")