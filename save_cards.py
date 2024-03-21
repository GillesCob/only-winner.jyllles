import imgkit

# Chemin vers votre fichier HTML
html_file = '/Users/gillescobigo/Documents/GitHub/only-winner.jyllles/cards.html'

# Chemin vers l'emplacement où vous souhaitez enregistrer l'image PNG
output_file = '/Users/gillescobigo/Desktop/test.png'

# Options pour imgkit (vous pouvez ajuster selon vos besoins)
options = {
    'quiet': '',
    'disable-smart-width': ''
}

# Convertir HTML en image PNG
imgkit.from_file(html_file, output_file, options=options)
