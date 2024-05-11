import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

# abrindo o navegador e logando no Extranet
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
navegador.implicitly_wait(10)
navegador.maximize_window()
navegador.get('https://extranet.lopesrio.com.br/crm/Listagem.aspx')
log = navegador.find_element(By.XPATH, '//input[@name="tLogin"]')
login = '15877017748'
for i in login:
    log.send_keys(i)
    time.sleep(0.01)
navegador.find_element(By.XPATH, '//input[@name="tSenha"]').send_keys('Lkm2020')
navegador.find_element(By.XPATH, '//a[@id="btLogin"]').click()

# criando listas vazias para armazenar valores
info_nome = []
info_atendido = []
info_nc = []
info_sc = []

# função que retorna o atributo "value" da lista de ofertas ativas:
def lista_ofertas():
    select_oferta = []
    selecionar_oferta = Select(navegador.find_element(By.XPATH, '//select[@id="ctl00_ContentPlaceHolder1_ddlOferta"]')).options
    for i in selecionar_oferta:
        valor = i.get_attribute('value')
        if not valor == '0':
            select_oferta.append(valor)
        
    return select_oferta

# função que retorna o nome das ofertas:
def nomes_ofertas():
    nome_oferta = Select(navegador.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlOferta"]')).options
    for i in nome_oferta:
        nome = i.text
        if not nome == 'SELECIONE':
            info_nome.append(nome)

# função que retorna os status da listagem (parâmetro e quantidade):
def status_listagem():
    status_list = []
    status = navegador.find_elements(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_pnlGerenciar"]/div[2]/div[1]/table/tbody/tr')
    for i in status:
        status_list.append(i)

    for i in range(1, len(status_list)):
        parametro = navegador.find_element(By.XPATH, f'//*[@id="ctl00_ContentPlaceHolder1_pnlGerenciar"]/div[2]/div[1]/table/tbody/tr[{i}]/td[1]').text
        qtd = navegador.find_element(By.XPATH, f'//*[@id="ctl00_ContentPlaceHolder1_pnlGerenciar"]/div[2]/div[1]/table/tbody/tr[{i}]/td[2]').text
        if parametro == 'ATENDIDO':
            info_atendido.append(qtd)
        elif parametro == 'NÃO CAPTURADO':
            info_nc.append(qtd)
        elif parametro == 'SEM CONTATO':
            info_sc.append(qtd)
        elif parametro == '':
            pass
    
    parametros = []
    
    for i in range(1, len(status_list)):
        texto_parametro = navegador.find_element(By.XPATH, f'//*[@id="ctl00_ContentPlaceHolder1_pnlGerenciar"]/div[2]/div[1]/table/tbody/tr[{i}]/td[1]').text
        if texto_parametro != '':
            parametros.append(texto_parametro)
        
    if 'ATENDIDO' not in parametros:
        info_atendido.append(0)

    if 'NÃO CAPTURADO' not in parametros:
        info_nc.append(0)

    if 'SEM CONTATO' not in parametros:
        info_sc.append(0)

# executando as funções:
nomes_ofertas()
offer_list = lista_ofertas()
for i in offer_list:
    navegador.find_element(By.XPATH, f'//option[@value="{i}"]').click()
    time.sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlListagem"]/option[2]').click()
    navegador.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btGerenciar"]').click()
    time.sleep(1)
    status_listagem()

# criando um dataframe contendo todas as informações
status_df = pd.DataFrame({
    'OFERTA': info_nome,
    'ATENDIDO': info_atendido,
    'NÃO CAPTURADO': info_nc,
    'SEM CONTATO': info_sc
})

# transformando o dataframe em um arquivo excel
data_hoje = datetime.now().strftime('%d-%m-%Y')

status_df.to_excel(f'status_listagens_{data_hoje}.xlsx', index=False)