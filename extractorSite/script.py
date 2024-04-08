import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import re
import time
import os

# Lecture du fichier Excel
df = pd.read_excel('C:\\Users\\Master\\Desktop\\Nomsite.xlsx')

# Initialisation du driver Selenium
edge_options = webdriver.EdgeOptions()
edge_options.add_argument("--headless")
service = Service('C:\\Users\\Master\\Documents\\edgedriver_win64\\msedgedriver.exe')
driver = webdriver.Edge(service=service, options=edge_options)

wait = WebDriverWait(driver, 10)

email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
phone_regex = r'(?:0[1-9](?:[\s.-]?\d{2}){4})|(?:\+33\s?(?:\(0\))?[1-9](?:[\s.-]?\d{2}){4})'

# Fonction pour trouver les informations de contact sur la page
def find_contact_info(page_content):
    emails = re.findall(email_regex, page_content)
    phones = re.findall(phone_regex, page_content)
    return emails[0] if emails else "Non trouvé", phones[0] if phones else "Non trouvé"

# Fonction pour naviguer vers la page de contact
def navigate_to_contact_page():
    try:
        contact_link = driver.find_element(By.XPATH, "//a[contains(text(), 'nous contacter') or contains(text(), 'Contact') or contains(text(), 'CONTACT') or contains(@href, 'contact')]")
        contact_link.click()
        time.sleep(3)  # Attendez que la page de contact se charge
        return driver.page_source
    except NoSuchElementException:
        return ""

# Liste des chaînes de caractères à exclure dans les URL
excluded_strings = [
    'pagesjaunes.fr',
    'abosociete.fr',
    'pappers.fr',
    'annuaire-entreprise.data.gouv.fr',
    'societe.com',
    'infonet.fr','linkedin.com', 'lefigaro', 'facebook', 'verif', 'kompass', 'europages', 'copainsdavant', 'pinterest', 'viadeo', 'instagram', 'wikipedia', 'manageo'
]

# Liste pour stocker les données collectées
data = []

for index, row in df.iterrows():
    company_name = row['NomSite']
    driver.get('https://www.google.com')
    
    try:
        search_box = wait.until(EC.visibility_of_element_located((By.NAME, 'q')))
        driver.execute_script("arguments[0].value = arguments[1];", search_box, company_name)
        driver.execute_script("arguments[0].form.submit();", search_box)

        # Attendre que les résultats de recherche soient visibles
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "h3")))

        # Récupérer tous les liens (éléments 'h3') des résultats de recherche
        results_links = driver.find_elements(By.CSS_SELECTOR, "h3")

        # Parcourir les résultats jusqu'à trouver un lien admissible
        for result_link in results_links:
            # Récupérer l'URL cible en vérifiant l'attribut href du parent <a>
            target_url = result_link.find_element(By.XPATH, "..").get_attribute("href")

            # Vérifier si l'URL ne contient aucune des chaînes exclues
            if not any(excluded_string in target_url for excluded_string in excluded_strings):
                # Si l'URL est admissible, cliquer sur le lien et sortir de la boucle
                driver.execute_script("arguments[0].click();", result_link)
                break
        else:
            # Si aucun lien admissible n'est trouvé, imprimer un message et passer à la prochaine itération
            print(f"Aucun lien admissible trouvé pour '{company_name}'. Passage à l'entreprise suivante.")
            continue

        time.sleep(5)  # Laissez du temps pour que la page du site web cible se charge

        current_url = driver.current_url
        print(f"URL pour {company_name}: {current_url}")
        
        page_content = driver.page_source
        email, phone = find_contact_info(page_content)
        
        if email == "Non trouvé" or phone == "Non trouvé":
            # Tentative de navigation vers la page de contact si l'un des éléments n'est pas trouvé
            page_content = navigate_to_contact_page()
            email, phone = find_contact_info(page_content)
        
        print(f"{company_name}: Email: {email}, Téléphone: {phone}, URL: {current_url}")
        
        # Ajouter les données collectées à la liste, y compris l'URL
        data.append({'NomSite': company_name, 'Email': email, 'Téléphone': phone, 'URL': current_url})

    except Exception as e:
        print(f"Une erreur s'est produite lors de la recherche pour {company_name}: {e}")
        continue

# Quitter le navigateur
driver.quit()

# Créer un DataFrame à partir de la liste des données collectées
df_updated = pd.DataFrame(data)

# Construire le chemin vers le bureau de l'utilisateur
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')

# Nom du fichier Excel à enregistrer
excel_file_name = 'Nomsite_mise_a_jour.xlsx'

# Chemin complet du fichier Excel
excel_file_path = os.path.join(desktop_path, excel_file_name)

# Enregistrer les données mises à jour dans un fichier Excel sur le bureau
df_updated.to_excel(excel_file_path, index=False)

print(f'Fichier enregistré sur le bureau : {excel_file_path}')
