from selenium import webdriver
import pandas as pd
import urllib.parse

if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)
    # Convertir le chemin du fichier en URL format file
    file_path = "/home/williamjamesmoriart/Automatic_test/DATA-TEST-CAP-IRVE-1.xlsx"
    file_url = urllib.parse.urljoin('file:', urllib.request.pathname2url(file_path))


    driver.get(file_url)
    data = pd.read_excel(file_path , sheet_name="Données")
# Itérer sur chaque ligne de données du DataFrame
    for index, row in data.iterrows():
        # Accéder aux données du projet
        projet_data = {
            "Nom_affaire": row["Nom_affaire"],
            "Date_Deadline": row["Date_Deadline"],
            "Adresse": row["Adresse"],
            "Prioritaire": row["Prioritaire"],
            # ... autres données du projet (accéder aux colonnes par leur nom)
        }
        print(projet_data["Nom_affaire"])
        # Remplir le formulaire web
        # ... code pour remplir les champs du formulaire avec les données du projet

        # Incrémenter le numéro de la ligne
   




    driver.quit()