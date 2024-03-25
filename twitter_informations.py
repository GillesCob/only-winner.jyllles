
import requests
import json

# Remplacez l'URL du webhook par l'URL de votre webhook Discord
webhook_url = 'https://discord.com/api/webhooks/1220361004479676456/ERaQqkUNyJgoJhYxpPRBznTX6LpF0M4as3E7IMuOa5Qhj7LrOuJoTyJ0v6RxqqMW-7sh'

# Liste des informations à envoyer au webhook
infos = ['Information 1', 'Information 2', 'Information 3']

# Boucle sur chaque information et envoie-la au webhook Discord
for info in infos:
    # Créez le message à envoyer
    message = {
        'content': info
    }
    # Envoyez le message au webhook Discord
    response = requests.post(webhook_url, data=json.dumps(message), headers={'Content-Type': 'application/json'})

    # Vérifiez si la requête a réussi
    if response.status_code == 204:
        print(f'Message "{info}" envoyé avec succès à Discord!')
    else:
        print(f'Échec de l\'envoi du message "{info}" à Discord.')
        print('Code de statut:', response.status_code)
