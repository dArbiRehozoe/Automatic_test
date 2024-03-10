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
    data = pd.read_excel(file_path)
# Itérer sur chaque ligne de données
    ligne_courante = 2
    for row in data.iter_rows(min_row=2):
        # Accéder aux données du projet
        projet_data = {
            "Nom_affaire": row[0].value,
            "Date_Deadline": row[1].value,
            "Adresse": row[2].value,
            "Prioritaire": row[3].value,
            # ... autres données du projet
        }
        print(projet_data["Nom_affaire"])
        # Remplir le formulaire web
        # ... code pour remplir les champs du formulaire avec les données du projet

        # Incrémenter le numéro de la ligne
        ligne_courante += 1




    driver.quit()