
#pip isntall openpyxl
#pip install selenium
#pip install urllib3
# pip install webdriver-manager

from selenium import webdriver #Abre o navegador e movimenta
from selenium.webdriver.chrome.service import Service 
from webdriver_manager.chrome import ChromeDriverManager # Acessar chrome driver na nuvem ( Sem instalar )
from selenium.webdriver.common.by import By #encontrar ids etc..
from selenium.webdriver.common.keys import Keys #Permite enviar as mensagens
import time #sleeps
import openpyxl #trabalhar com planilhas
import urllib #formatador da mensagem

###### Tira os erros de log no terminal
options = webdriver.ChromeOptions()
options.add_argument("--disable-logging")
options.add_argument("--log-level=3")
######

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
driver.maximize_window()
driver.get("https://web.whatsapp.com/")
time.sleep(20)
contatos = openpyxl.load_workbook("hubla.xlsx")
sheet = contatos.active
m_row = sheet.max_row

for i in range(2, m_row + 1):
    name = sheet.cell(row = i, column = 3).value
    number = sheet.cell(row = i, column = 4).value

    mensagem =f"Ola, {name}, com o numero: {number}, voce foi selecionado com exito pelo robo!"
    texto = urllib.parse.quote(mensagem)
    print(texto)
    link = f"https://web.whatsapp.com/send?phone={number}&text={texto}"
    driver.get(link)
    print("Entrou no link..")
    time.sleep(10)

    campo = driver.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p")
    campo.send_keys(Keys.ENTER)

    time.sleep(10)
    print('esperando 10s ..')






"""""

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
driver.maximize_window()
# driver.get("https://web.whatsapp.com/")
# time.sleep(20)
contatos = openpyxl.load_workbook("hubla.xlsx")
sheet = contatos.active
m_row = sheet.max_row

for i in range(2, m_row + 1):
    name = sheet.cell(row = i, column = 3).value
    number = sheet.cell(row = i, column = 4).value

    mensagem =f"Ola, {name}, com o numero: {number}, voce foi selecionado com exito pelo robo!"
    texto = urllib.parse.quote(mensagem)
    print(texto)
    link = f"https://web.whatsapp.com/send?phone={number}&text={texto}"
    print(link)
    driver.get(link)
    # print("Entrou no link..")
    # time.sleep(30)
    while len(driver.find_elements(By.ID, "side")) < 1:
        time.sleep(1)
    time.sleep(2)

    if len(driver.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div')) < 1:
        # time.sleep(10)
        campo = driver.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p")
        campo.send_keys(Keys.ENTER)

    time.sleep(10)
    print('esperando 10s ..')


"""