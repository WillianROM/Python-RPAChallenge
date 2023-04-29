####### Importação para ler a planilha
import pandas as pd

####### Importações para automatizar a web
from selenium import webdriver # Navegador
from selenium.webdriver.common.by import By  # Achar os elementos
from selenium.webdriver.common.keys import Keys # Para digitar no teclado

####### Gerenciador de navegador
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

import time

# Ler dados da planilha
nome_do_arquivo = "challenge.xlsx"
df = pd.read_excel(nome_do_arquivo)
#print(df)

# Abrir o navegador Chrome
servico = ChromeService(ChromeDriverManager().install())  #Verifica se a versão do Chrome está atualizado
driver = webdriver.Chrome(service=servico) #Abre o chrome com a versão atualizada
driver.implicitly_wait(5) #Tempo de espera implicita
driver.maximize_window() #Deixar o navegador com a tela máxima

#Abrir o site RPAChallenge
driver.get("https://www.rpachallenge.com/")


# Clique no botão Start
BUTTON_START = driver.find_element(By.XPATH,"//button[contains(text(),'Start')]")
BUTTON_START.click()


for index, row in df.iterrows():
    #Mapeamento dos componentes da página
    INPUT_ADDRESS = driver.find_element(By.XPATH,"//label[contains(text(),'Address')]/following-sibling::input")
    INPUT_COMPANY_NAME = driver.find_element(By.XPATH,"//label[contains(text(),'Company Name')]/following-sibling::input")
    INPUT_FIRST_NAME = driver.find_element(By.XPATH,"//label[contains(text(),'First Name')]/following-sibling::input")
    INPUT_LAST_NAME = driver.find_element(By.XPATH,"//label[contains(text(),'Last Name')]/following-sibling::input")
    INPUT_EMAIL = driver.find_element(By.XPATH,"//label[contains(text(),'Email')]/following-sibling::input")
    INPUT_PHONE_NUMBER = driver.find_element(By.XPATH,"//label[contains(text(),'Phone Number')]/following-sibling::input")
    INPUT_ROLE_IN_COMPANY = driver.find_element(By.XPATH,"//label[contains(text(),'Role in Company')]/following-sibling::input")
    INPUT_SUBMIT = driver.find_element(By.XPATH,"//input[@type='submit']")
    
    #Preenchimento dos componentes da página
    INPUT_FIRST_NAME.send_keys(row["First Name"])
    INPUT_LAST_NAME.send_keys(row["Last Name"])
    INPUT_COMPANY_NAME.send_keys(row["Company Name"])
    INPUT_ROLE_IN_COMPANY.send_keys(row["Role in Company"])
    INPUT_ADDRESS.send_keys(row["Address"])
    INPUT_EMAIL.send_keys(row["Email"])
    INPUT_PHONE_NUMBER.send_keys(row["Phone Number"])
    
    #Clique no botão Submit
    INPUT_SUBMIT.click()

time.sleep(60)
driver.close