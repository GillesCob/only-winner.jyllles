from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image

# Chemin du fichier HTML local
file_path = "chemin_vers_votre_fichier/index.html"

# Configuration du navigateur
options = webdriver.ChromeOptions()
options.add_argument('headless')  # Exécuter le navigateur en mode headless (sans interface graphique)

# Initialisation du navigateur
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

# Charger le fichier HTML local
driver.get("file:///" + file_path)

# Attendre quelques secondes pour s'assurer que tous les éléments sont chargés
driver.implicitly_wait(5)

# Prendre une capture d'écran de la page entière
screenshot = driver.get_screenshot_as_png()

# Fermer le navigateur
driver.quit()

# Enregistrer la capture d'écran en tant qu'image PNG
output_path = "screenshot.png"
with open(output_path, "wb") as f:
    f.write(screenshot)

print(f"Capture d'écran enregistrée sous : {output_path}")
