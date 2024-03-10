from selenium import webdriver
import pandas as pd
import urllib.parse

if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)
    # Convertir le chemin du fichier en URL format file
    file_path = "/home/fesia/testSelenium/selenium_env/DATA TEST CAP-IRVE (1).xlsx"
    file_url = urllib.parse.urljoin('file:', urllib.request.pathname2url(file_path))
    driver.get(file_url)
    data = pd.read_excel(file_path)
    print(data)




    driver.quit()