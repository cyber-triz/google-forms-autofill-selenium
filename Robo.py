import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Caminho para o arquivo Excel
excel_file = 'teste.xlsx'

# Inicializar o serviço do ChromeDriver
chrome_driver_path = 'C:\Pasta4\Contatos\chromedriver_win32\chromedriver.exe'  # Insira o caminho para o ChromeDriver
service = Service(chrome_driver_path)

# Opções do ChromeDriver
options = Options()

# Inicializar o driver do Chrome
driver = webdriver.Chrome(service=service, options=options)

# Carregar o formulário do Google Forms
form_url = 'https://docs.google.com/forms/d/e/1FAIpQLSdRjCf8S_GgUvf51afG_HwJUbSdICYYgoPeJZG_FuuRm8SYUw/viewform?usp=sf_link'  # Insira a URL do seu Google Forms
driver.get(form_url)

# Abrir o arquivo Excel
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Loop pelos dados do Excel e enviar para o Google Forms
for row in sheet.iter_rows(min_row=2, values_only=True):
    # Preencher o formulário com os dados da linha atual
    campo1 = str(row)  # Supondo que há três campos no formulário web

    campo1_input = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/input')  # Insira o XPath do campo 1
    campo1_input.send_keys(campo1)

    

    # Enviar o formulário
    submit_button = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span')  # Insira o XPath do botão de enviar
    submit_button.click()

    submit_button = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[1]/div/div[4]/a')  # Insira o XPath do botão de enviar
    submit_button.click()

# Fechar o navegador e o arquivo Excel
driver.quit()
workbook.close()
