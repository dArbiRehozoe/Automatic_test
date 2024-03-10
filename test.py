from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
import urllib.parse

options = webdriver.ChromeOptions()
options.add_argument("--headless")
driver = webdriver.Chrome(options=options)
    # Convertir le chemin du fichier en URL format file
file_path = "/home/williamjamesmoriart/Automatic_test/DATA-TEST-CAP-IRVE-1.xlsx"
file_url = urllib.parse.urljoin('file:', urllib.request.pathname2url(file_path))


driver.get(file_url)
data = pd.read_excel(file_path)
# Itérer sur chaque ligne de données du DataFrame

# Configurer les options du navigateur
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# Initialiser le navigateur
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://cap-irve-stage.digetit.com/")

# Remplir le formulaire de connexion
email_input = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "#email"))
) 
email_input.send_keys("natifesia@gmail.com")

password_input = driver.find_element(By.ID, "password")
password_input.send_keys(" natifesia-test@cap-irve-dev")

login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
login_button.click()

## Aller sur la page "Nouveau Projet" avec une attente explicite
wait = WebDriverWait(driver, 10)
nouveau_projet_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "css-1gactu1")))
nouveau_projet_button.click()



# Récupérer les données du premier projet à partir du DataFrame


## Simuler le clic sur le "combobox" pour ouvrir la liste déroulante
combobox = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "demo-simple-select"))
)
combobox.click()

# Sélectionner une option en cliquant sur l'élément li correspondant à l'option spécifique
option_to_select = "test"  # Remplacez par le texte de l'option que vous souhaitez sélectionner
option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, f"//li[@role='option'][text(🙁'{option_to_select}']"))
)
option.click()


# ... Continuer à remplir les autres champs avec les valeurs extraites du fichier Excel
for index, row in data.iterrows():
        # Accéder aux données du projet
    projet_data = {
            "client": row["client"],
            "type equipement":row["type equipement"]

            # ... autres données du projet (accéder aux colonnes par leur nom)
        }
    print(projet_data["client"])
    print(projet_data["type equipement"])
    # Remplir le formulaire web

    nom_affaire_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[name="client"]'))
    )
    nom_affaire_input.send_keys(projet_data["client"])

    nom_affaire_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[name="Nom_affaire"]'))
    )
    nom_affaire_input.send_keys("entreprise clairina")

    # Remplir le champ "Date de Deadline"
    date_deadline_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[name="Date_Deadline"]'))
    )
    date_deadline_input.clear()  # Effacer toute date existante
    date_deadline_input.send_keys("2024-03-15")  # Remplacer par la date que vous souhaitez saisir (au format "YYYY-MM-DD")
    # Remplir le champ "Addresse"
    adresse_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[name="adresse"]'))
    )
    adresse_input.send_keys("Madagascar")

    # Sélectionner une option dans le "combobox" pour "Prioritaire"
    check_priotitaire = driver .find_element(By .ID, "mui-component-select-prioritaire")
    check_priotitaire.click()

    webdriver.ActionChains(driver).send_keys(Keys.ARROW_UP).perform()
    webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()


    submit_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/main/div[2]/div/div/div[2]/div/form/div/div[8]/div/button[2]')
    submit_button.click()

    #on commence par l' etape 2
    # Remplir le champ "Type d'abonnement"
    type_abonnement_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-typeAbonnement"))
    )
    type_abonnement_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner un élément spécifique dans la liste déroulante
    specific_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'li[data-value="Vert"]'))  # Utilisation du sélecteur CSS pour l'élément spécifique
    )
    specific_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Remplir le champ "Régime"
    regime_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-regime"))
    )
    regime_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner un élément spécifique dans la liste déroulante
    specific_regime_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'li[data-value="TT"]'))  # Utilisation du sélecteur CSS pour l'élément spécifique
    )
    specific_regime_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Remplir le champ "Puissance Souscrite"
    puissance_souscrite_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, ":r1c:"))
    )
    puissance_souscrite_input.clear()  # Effacer toute valeur existante
    puissance_souscrite_input.send_keys("2")  # Remplacer "VotreValeur" par la valeur que vous choisi

    # Remplir le champ "Unité de Puissance Souscrite"
    unite_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-unite_puissance_soucrite"))
    )
    unite_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner l'unité spécifique dans la liste déroulante
    specific_unite_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'li[data-value="KVA"]'))  # Utilisation du sélecteur CSS pour l'unité spécifique
    )
    specific_unite_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Remplir le champ "Puissance Déjà Consommée"
    puissance_consommee_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, ":r11:"))
    )
    puissance_consommee_input.clear()  # Effacer toute valeur existante
    puissance_consommee_input.send_keys("2")  # Remplacer "VotreValeur" par la valeur que vous souhaitez 

    # Remplir le champ "Unité de Puissance Consommée"
    unite_consommee_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-unite_puissance_consommee"))
    )
    unite_consommee_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner l'unité spécifique dans la liste déroulante
    specific_unite_consommee_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'li[data-value="KW"]'))  # Utilisation du sélecteur CSS pour l'unité spécifique
    )
    specific_unite_consommee_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Remplir le champ "Mode de Pose vers TGBT"
    mode_pose_tgbt_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-mode_pose_tgbt"))
    )
    mode_pose_tgbt_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner le mode spécifique dans la liste déroulante
    specific_mode_pose_tgbt_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'li[data-value="13"]'))  # Utilisation du sélecteur CSS pour le mode spécifique
    )
    specific_mode_pose_tgbt_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Remplir le champ "Distance entre la source et le TGBT"
    distance_tgbt_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, ":r14:"))
    )
    distance_tgbt_input.clear()  # Effacer toute valeur existante
    distance_tgbt_input.send_keys("4")

    # Sélectionner "Oui" dans le champ "Souhaitez-vous forcer la protection"
    forcing_protection_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-equipement_existant_tgt"))
    )
    forcing_protection_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner "Oui" dans la liste déroulante
    specific_forcing_protection_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//li[text(🙁"Oui"]'))  # Utilisation du sélecteur XPath pour l'élément spécifique
    )
    specific_forcing_protection_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Sélectionner le constructeur dans le champ "Constructeur"
    constructeur_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-Constructeur"))
    )
    constructeur_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner un élément spécifique dans la liste déroulante
    specific_constructeur_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//li[text(🙁"LEGRAND"]'))  # Remplacer "Legrand" par le constructeur spécifique que vous souhaitez choisir
    )
    specific_constructeur_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Sélectionner le type d'équipement dans le champ "Type d'équipement"
    type_equipement_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-type_equipement"))
    )
    type_equipement_dropdown.click()  # Pour ouvrir la liste déroulante

    # Sélectionner un type d'équipement spécifique dans la liste déroulante
    specific_type_equipement_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//li[text(🙁"Disjonct. D"]'))  # Remplacer "VotreTypeEquipement" par le type d'équipement spécifique que vous souhaitez choisir
    )
    specific_type_equipement_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Ouvrir la liste déroulante du nom de l'équipement
    nom_equipement_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-nom_equipement_tgbt"))
    )
    nom_equipement_dropdown.click()

    # Sélectionner "NS2000H 2000A 4P4D" dans la liste déroulante du nom de l'équipement
    specific_nom_equipement_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//li[text(🙁"NS2000H 2000A 4P4D"]'))  # Utiliser XPath pour sélectionner l'option "NS2000H 2000A 4P4D"
    )
    specific_nom_equipement_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Remplir le champ "Valeur de IK1"
    ik1_valeur_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, ":r19:"))
    )
    ik1_valeur_input.clear()  # Effacer toute valeur existante
    ik1_valeur_input.send_keys("2")  # Remplacer "02" par la valeur que vous souhaitez saisir

    # Remplir le champ "Valeur de LK3"
    lk3_valeur_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, ":r1a:"))
    )
    lk3_valeur_input.clear()  # Efface toute valeur existante
    lk3_valeur_input.send_keys("2")  # Remplacez "VotreValeurLK3" par la valeur que vous souhaitez saisir


    # Ouvrir la liste déroulante
    td_irve_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mui-component-select-td_irve"))
    )
    td_irve_dropdown.click()

    # Sélectionner "Oui" dans la liste déroulante
    specific_td_irve_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'li[data-value="true"]'))  # Utiliser le sélecteur CSS pour sélectionner "Oui" basé sur l'attribut "data-value"
    )
    specific_td_irve_element.click()

    # Attendre un bref délai pour permettre à l'élément d'être sélectionné
    time.sleep(1)

    # Cliquer sur le bouton "Suivant"
    suivant_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.MuiButton-root[type="submit"]'))  # Utilisation du sélecteur CSS pour le bouton "Suivant"
    )
    suivant_button.click()

    # Attendre un bref délai pour permettre à la nouvelle page de se charger
    time.sleep(1)