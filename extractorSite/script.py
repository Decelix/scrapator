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

def find_contact_info(page_content):
    emails = re.findall(email_regex, page_content)
    filtered_emails = [email for email in emails if not any(ext in email for ext in ['.jpg', '.jpeg', '.png', '.webp', '.svg'])]
    phones = re.findall(phone_regex, page_content)
    return filtered_emails[0] if filtered_emails else "Non trouvé", phones[0] if phones else "Non trouvé"

def navigate_to_contact_page():
    try:
        contact_link = driver.find_element(By.XPATH, "//a[contains(text(), 'nous contacter') or contains(text(), 'Contact') or contains(text(), 'CONTACT') or contains(@href, 'contact')]")
        contact_link.click()
        time.sleep(3)
        return driver.page_source
    except NoSuchElementException:
        return ""

excluded_strings = [
    'pagesjaunes.fr', 'abosociete.fr', 'pappers.fr', 'annuaire-entreprises.data.gouv.fr', 'societe.com', 'infonet.fr', 'linkedin.com', 'lefigaro', 'facebook', 'verif', 'kompass', 'europages', 'copainsdavant', 'pinterest', 'viadeo', 'instagram', 'wikipedia', 'manageo', 'xerfi.com'
]

data = []

for index, row in df.iterrows():
    company_name = row['NomSite']
    # Initialisation du dictionnaire avec des valeurs par défaut
    company_info = {'NomSite': company_name, 'Email': "Non trouvé", 'Téléphone': "Non trouvé", 'URL': "Non trouvé"}

    try:
        driver.get('https://www.google.com')
        search_box = wait.until(EC.visibility_of_element_located((By.NAME, 'q')))
        driver.execute_script("arguments[0].value = arguments[1];", search_box, company_name)
        driver.execute_script("arguments[0].form.submit();", search_box)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "h3")))
        results_links = driver.find_elements(By.CSS_SELECTOR, "h3")
        
        for result_link in results_links:
            target_url = result_link.find_element(By.XPATH, "..").get_attribute("href")
            if not any(excluded_string in target_url for excluded_string in excluded_strings):
                driver.execute_script("arguments[0].click();", result_link)
                break
        else:
            print(f"Aucun lien admissible trouvé pour '{company_name}'. Passage à l'entreprise suivante.")
            data.append(company_info)
            continue

        time.sleep(5)
        current_url = driver.current_url
        page_content = driver.page_source
        email, phone = find_contact_info(page_content)
        
        if email == "Non trouvé" or phone == "Non trouvé":
            page_content = navigate_to_contact_page()
            email, phone = find_contact_info(page_content)
        
        company_info.update({'Email': email, 'Téléphone': phone, 'URL': current_url})

    except Exception as e:
        print(f"Une erreur s'est produite lors de la recherche pour {company_name}: {e}")
    
    finally:
        data.append(company_info)

driver.quit()
df_updated = pd.DataFrame(data)
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
excel_file_name = 'Nomsite_mise_a_jour.xlsx'
excel_file_path = os.path.join(desktop_path, excel_file_name)
df_updated.to_excel(excel_file_path, index=False)
print(f'Fichier enregistré sur le bureau : {excel_file_path}')

