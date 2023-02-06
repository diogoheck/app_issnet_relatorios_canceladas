from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
from selenium import webdriver
from exception.lancar_excecao import lancamento_excecao
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
PASTA_XML = r'C:\ISS\xml_contratados'
# PASTA_XML = r'C:\xml'


def pega_valores_botao(elemento):
    lista = []
    valores = elemento.get_attribute("value").split('-')
    lista.append(int(valores[0].strip()))
    lista.append(int(valores[1].strip()))
    return lista


def pega_botoes(driver):
    dicionario = {}
    for num in range(5):
        botao = 'btn'+str(num+1)
        elementos = driver.find_element(By.ID, botao)
        tupla = tuple(pega_valores_botao(elementos))
        dicionario.update({tupla: botao})
    return dicionario


def login_site(driver):
    driver.get(
        'https://www.issnetonline.com.br/santamaria/online/login/login.aspx')


def inserir_cfp(driver, CPF):
    driver.find_element(By.XPATH,
                        '//*[@id="txtLogin"]').send_keys(CPF)


def clicar_botao_senha(driver, botao):
    driver.find_element(By.ID, botao[0]).click()


def clicar_botao_acesso(driver):
    driver.find_element(By.XPATH, '//*[@id="btnAcessar"]').click()


def clicar_msg_expiracao_senha(driver):
    driver.find_element(By.XPATH,
                        '/html/body/div[1]/div/div/div[3]/div/button').click()


def inserir_senha(driver, CPF, senhas):

    lancamento_excecao(login_site, driver)

    driver.maximize_window()

    time.sleep(10)

    dicionario = pega_botoes(driver)

    lancamento_excecao(inserir_cfp, driver, CPF)

    for senha in senhas:
        botao = [botao for k, botao in dicionario.items() if senha in k]
        lancamento_excecao(clicar_botao_senha, driver, botao)

    lancamento_excecao(clicar_botao_acesso, driver)

    # para mensagem de expiração de senha
    lancamento_excecao(clicar_msg_expiracao_senha, driver, segundos=1)


def criar_conexao(CPF, senhas):
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": PASTA_XML,
        "download.prompt_for_download": False,
        # Desabilitar a confirmação de download
        'download.prompt_for_download': False,
        # Desabilitar notificações
        'profile.default_content_setting_values.notifications': 2,
        # Permitir multiplos downloads
        'profile.default_content_setting_values.automatic_downloads': 1,

    }

    options.add_experimental_option("prefs", prefs)

    arguments = ['--lang=pt-BR', '--window-size=800,600',
                 '--disable-web-security']
    for argument in arguments:
        options.add_argument(argument)

    driver = uc.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        chrome_options=options)
    # driver = uc.Chrome()

    inserir_senha(driver, CPF, senhas)
    return driver
