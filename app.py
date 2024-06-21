import openpyxl
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Inicializa o driver
driver = webdriver.Chrome()

# Abre o WhatsApp Web
driver.get('https://web.whatsapp.com/')
time.sleep(15)  # Aguardar manualmente o login no WhatsApp Web

# Carrega a planilha de clientes
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Plan1']

# Itera sobre as linhas da planilha
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = f'Olá {nome}, é apenas um teste do joãozinho'
    link_mens_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    
    # Abre o link de mensagem no WhatsApp
    driver.get(link_mens_whats)
    time.sleep(10)  # Aguarda a página carregar

    try:
        # Aguarda o botão de enviar aparecer e clica
        seta = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span'))
        )
        seta.click()
        time.sleep(5)  # Aguarda um pouco após enviar

    except Exception as e:
        print(f'Não foi possível enviar a mensagem para {nome}: {str(e)}')
        with open('erros.csv', 'a', newline='', encoding='UTF-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}\n')

# Fecha o driver
driver.quit()
