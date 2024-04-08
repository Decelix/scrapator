from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

def extract_data(url):
    edge_driver_path = 'C:\\Users\\Master\\Documents\\edgedriver_win64\\msedgedriver.exe'
    service = Service(edge_driver_path)
    options = webdriver.EdgeOptions()
    # Activer le mode sans tête
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')  # Parfois nécessaire selon la version et l'environnement

    driver = webdriver.Edge(service=service, options=options)

    try:
        driver.get(url)

        WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".result-item .title span"))
        )

        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')

        file_path = 'C:\\Users\\Master\\Desktop\\ExtractNomSite\\resultats.txt'
        with open(file_path, 'a', encoding='utf-8') as file:
            titles = soup.select(".result-item .title span")
            for title in titles:
                file.write(title.text.strip() + '\n')
    finally:
        driver.quit()

# URL de départ
base_url = "https://annuaire-entreprises.data.gouv.fr/rechercher?terme=&etat=A&naf=13.20Z&page="
page_number = 1

while True:
    url = base_url + str(page_number)
    print("Extraction des données depuis :", url)
    try:
        extract_data(url)
    except Exception as e:
        print("URL incorrecte ou fin de l'extraction. Erreur:", e)
        break
    else:
        page_number += 1
