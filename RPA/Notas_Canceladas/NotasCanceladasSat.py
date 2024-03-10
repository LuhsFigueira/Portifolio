from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
import requests
import time
from anticaptchaofficial.imagecaptcha import *
from selenium.webdriver.firefox.options import Options
import os
import shutil  # Para renomear/mover o arquivo


#Função que trata a quebra do captacha

def resolver_captcha ():
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_rdlSituacao_2").click()
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField").send_keys('85778074000106')
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_datDataInicial").send_keys('30112023')
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_datDataFinal").send_keys('08022024')
    time.sleep(2)
    captcha_element = driver.find_element(By.CSS_SELECTOR, '#Body_Main_Main_sepConsultaNfpe_ctl17 > img:nth-child(1)')
    # Gera a screenshot do elemento 
    captcha_element.screenshot('captacha.png')

    #Faz a trativa do captacha
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key("chave_acesso_quebra_captcha")

    # Specify softId to earn 10% commission with your app.
    # Get your softId here: https://anti-captcha.com/clients/tools/devcenter
    solver.set_soft_id(0)

    captcha_text = solver.solve_and_return_solution(r'captacha.png')
    if captcha_text != 0:
        print ("captcha text "+captcha_text)
    else:
        print ("task finished with error "+solver.error_code)

    time.sleep(15)
    element = driver.find_element(By.CSS_SELECTOR, "div.input-group:nth-child(1) > input:nth-child(1)").send_keys(captcha_text)
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_btnBuscar").click()

#Função para gerar relatório

def gera_relatorio():

    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_grpResultado_actConfiguration").click()
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#btn0").click()
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_grpResultado_ColumnConfigurationWindow_ctl12_6").click()
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_grpResultado_ColumnConfigurationWindow_uppPanelButtons > a:nth-child(1)").click()
    time.sleep(2)
    element = driver.find_element(By.CSS_SELECTOR, ".btn-toolbar > div:nth-child(1)").click()
    time.sleep(15)
    # Certifique-se de esperar o suficiente para o download ser concluído antes de tentar renomear o arquivo
        # Espera pelo download do arquivo
    tempo_espera_maximo = 120  # Ajuste conforme a necessidade
    tempo_inicial = time.time()
    while True:
        arquivo_baixado = max([os.path.join(diretorio_download, f) for f in os.listdir(diretorio_download)], key=os.path.getctime)
        if ".part" not in arquivo_baixado:  # Verifica se o download foi concluído
            break
        elif time.time() - tempo_inicial > tempo_espera_maximo:
            print("Tempo de espera pelo download excedido.")
            break
        time.sleep(1)  # Espera antes de verificar novamente

    # Renomeia o arquivo baixado para o nome específico
    shutil.move(arquivo_baixado, caminho_arquivo_final)

def envio_email_erro ():
    from pykeepass import PyKeePass
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    import smtplib
    from email.mime.application import MIMEApplication
    from openpyxl.utils import get_column_letter
    # Criar corpo do e-mail em HTML
    email_body = """
    <html>
    <head></head>
    <body>
    <p>Olá, </p>
    <p>Houveram {} tentativas de quebra de captcha sem sucesso! Verifique o processo! </p>
    </body>
    </html>
    """

    # Incluir o DataFrame ocultando algumas colunas no corpo do e-mail HTML
    html = email_body.format(max_tentativas)

    # Configurar e-mail
    msg = MIMEMultipart()
    msg['From'] = 'dados@riosulense.com.br'
    msg['To'] = 'luis.oliveira@riosulense.com.br'
    msg['Subject'] = 'Erro na quebra de captcha - Automatização S.'

    # Adicionar o corpo do e-mail em HTML
    msg.attach(MIMEText(html, 'html'))

    # Enviar e-mail
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'dados@riosulense.com.br'
    smtp_password = 'password'

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(msg)




#Caminho e nome do arquivo para fazer download
diretorio_download = r'C:\Users\luis.oliveira\Desktop\CapturadorXML-2\Envio  Relatório NF'
nome_arquivo_final = 'NotaCanceladaSAT.xlsx'
caminho_arquivo_final = os.path.join(diretorio_download, nome_arquivo_final)


# Cria uma instância de Options
options = Options()

# Configurações personalizadas
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.dir", diretorio_download)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# Inicializa o driver com as opções definidas
driver = webdriver.Firefox(options=options)
driver.get("https://encurtador.com.br/gAM02")
element = driver.find_element(By.CSS_SELECTOR, "#Body_pnlMain_tbxUsername").send_keys('usuario')
time.sleep(2)
element = driver.find_element(By.CSS_SELECTOR, "#Body_pnlMain_tbxUserPassword").send_keys('password')
time.sleep(2)
element = driver.find_element(By.CSS_SELECTOR, "#Body_pnlMain_btnLogin").click()
time.sleep(2)
element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_rdlSituacao_2").click()
time.sleep(2)
element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField").send_keys('85778074000106')
time.sleep(2)
element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_datDataInicial").send_keys('30112023')
time.sleep(2)
element = driver.find_element(By.CSS_SELECTOR, "#Body_Main_Main_sepConsultaNfpe_datDataFinal").send_keys('08022024')
time.sleep(2)
resolver_captcha()

from selenium.common.exceptions import NoSuchElementException
import time

tentativas = 0
max_tentativas = 3  # Define um limite de tentativas para evitar loop infinito

while tentativas < max_tentativas:
    try:
        # Tenta encontrar o elemento
        erro = driver.find_element(By.CSS_SELECTOR, ".sat-vs-error")
        resolver_captcha()

    except NoSuchElementException:
        gera_relatorio()
        break  # Sai do loop se não houver erro
    tentativas += 1

else: 
    envio_email_erro()







    









