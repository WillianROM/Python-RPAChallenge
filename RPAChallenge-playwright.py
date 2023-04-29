####### Importação para ler a planilha
import pandas as pd

####### Importações para automatizar a web
from playwright.sync_api import sync_playwright

import time

# Ler dados da planilha
nome_do_arquivo = "challenge.xlsx"
df = pd.read_excel(nome_do_arquivo)
#print(df)

with sync_playwright() as p:
    # Abrir o navegador Chrome
    driver = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = driver.new_context(no_viewport=True)
        
    page = context.new_page()

    #Abrir o site RPAChallenge
    page.goto("https://www.rpachallenge.com/")
   
    # Clique no botão Start
    BUTTON_START = page.locator("xpath=//button[contains(text(),'Start')]")
    BUTTON_START.click()


    for index, row in df.iterrows():
        #Mapeamento dos componentes da página
        INPUT_ADDRESS = page.locator("xpath=//label[contains(text(),'Address')]/following-sibling::input")
        INPUT_COMPANY_NAME = page.locator("xpath=//label[contains(text(),'Company Name')]/following-sibling::input")
        INPUT_FIRST_NAME = page.locator("xpath=//label[contains(text(),'First Name')]/following-sibling::input")
        INPUT_LAST_NAME = page.locator("xpath=//label[contains(text(),'Last Name')]/following-sibling::input")
        INPUT_EMAIL = page.locator("xpath=//label[contains(text(),'Email')]/following-sibling::input")
        INPUT_PHONE_NUMBER = page.locator("xpath=//label[contains(text(),'Phone Number')]/following-sibling::input")
        INPUT_ROLE_IN_COMPANY = page.locator("xpath=//label[contains(text(),'Role in Company')]/following-sibling::input")
        INPUT_SUBMIT = page.locator("xpath=//input[@type='submit']")
        
        #Preenchimento dos componentes da página
        INPUT_FIRST_NAME.fill(row["First Name"])
        INPUT_LAST_NAME.fill(row["Last Name"])
        INPUT_COMPANY_NAME.fill(row["Company Name"])
        INPUT_ROLE_IN_COMPANY.fill(row["Role in Company"])
        INPUT_ADDRESS.fill(row["Address"])
        INPUT_EMAIL.fill(row["Email"])
        INPUT_PHONE_NUMBER.fill(str(row["Phone Number"]))
        
        #Clique no botão Submit
        INPUT_SUBMIT.click()
    
    time.sleep(60)   
    driver.close