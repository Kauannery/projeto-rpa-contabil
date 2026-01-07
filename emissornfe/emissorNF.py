#!/usr/bin/env python
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import datetime
import os
import time

############################Configurações##################################
#Matricula do EmbraSystem
dsEmbraMatricula = ''
#Senha do EmbraSystem
dsEmbraPass = ''
#CNPJ Usado no EmissorNacional
dsEmissorCNPJ = ''
#Senha do EmissorNacional
dsEmissorPass = ''
#Valor da nota a ser emitida incluindo centavos sem virgula
vlNota = '1500'
#Descrição de busca de Município de forma a aparecer nos resultados do filtro
dsBuscaMunicipio = ''
#nome do Município exatamente como aparece no resiltado do filtro
dsMunicipio = ''
#CNPJ do tomador
cdTomador = ''
#Descrição do código tributário de forma a ser o único no resultado do filtro
dsTributario = ''
#Descrição do Serviço conforme cadastrado no EmbraSystem
dsServico = ''
#Pasta raiz onde as notas serão salvas. Diretamente abaixo criará a estrutura YYYY/mm/dd/
dsDownloadPath = ''
#Indica se rodará em um terminal sem interface gráfica
inTerminal = False
###########################################################################
#recupera a data atual para ser usada no preenchimento da nota e na configuração da pasta de Download do Driver
data_atual = datetime.date.today()

#criando pasta de Download
ano = str(data_atual.year)
mes = str(data_atual.month).zfill(2)
dsCompleteDownloadPath = os.path.join(dsDownloadPath, ano, mes)

os.makedirs(dsCompleteDownloadPath, exist_ok=True)
print("as notas serão salvas em:"+dsCompleteDownloadPath)
#instanciando o Chrome abrindo o EmissorNacional
prefs = {
    "download.default_directory": dsCompleteDownloadPath
}

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--safebrowsing-disable-download-protection')
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument('--window-size=1920,1080')

if inTerminal:
    chrome_options.add_argument('--headless=chrome')

#iniciando o webdiver
navegador = webdriver.Chrome(options=chrome_options)

#iniciando a emissão de nota fiscal
navegador.get('https://www.nfse.gov.br/EmissorNacional/Login')
wait = WebDriverWait(navegador, 10)

#aceitando site inseguro
print("aceitando site inseguro...")
try:
    navegador.find_element(By.XPATH, '//*[@id="details-button"]')
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="details-button"]'))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="proceed-link"]'))).click()
except Exception as e:
    print("Não havia página de segurança.")

try:
    #logando no EmissorNacional
    print("logando no EmissorNacional...")
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Inscricao"]'))).send_keys(dsEmissorCNPJ)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Senha"]'))).send_keys(dsEmissorPass)
    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button'))).click()

    #time.sleep(300)
    #iniciando emissão de nova nota
    print("iniciando emissão de nova nota...")
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wgtAcessoRapido"]/div[2]/a[1]/img'))).click()

    #calculando e preenchendo data de competência
    print("calculando e preenchendo data de competência...")
    primeiro_dia_do_mes = data_atual.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_do_mes - datetime.timedelta(days=1)
    data_formatada = ultimo_dia_mes_anterior.strftime("%d/%m/%Y")
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="DataCompetencia"]'))).send_keys(data_formatada)
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="DataCompetencia"]'))).send_keys(Keys.TAB)

    #selecionando tomador baseado no CNPJ da Embrasac
    print("selecionando tomador baseado no CNPJ da Embrasac...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnMaisInfoEmitente"]')))
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="pnlTomador"]/div[1]/div/div/div[2]/label/span/i'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_Tomador_Inscricao_historico"]'))).click()
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), cdTomador )]"))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="btnImportar"]'))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="btn_Tomador_EnderecoNacional_CEP"]'))).click()
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="btnAvancar"]'))).click()

    #selecionando Município de prestação do Serviço
    print("selecionando Município de prestação do Serviço...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlLocalPrestacao"]/div/div/div[2]/div/span[1]/span[1]/span'))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(dsBuscaMunicipio)
    wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '"+dsMunicipio+"')]"))).click()

    #selecionando código de tributação nacional
    print("selecionando código de tributação nacional...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlServicoPrestado"]/div/div[1]/div/div/span[1]/span[1]/span'))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(dsTributario)
    wait.until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="select2-ServicoPrestado_CodigoTributacaoNacional-results"]/li'),dsTributario))
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(Keys.TAB)

    #selecionando "Não" para a opção: O serviço prestado é um caso de: exportação, imunidade ou não incidência do ISSQN?*
    print('selecionando "Não" para a opção: O serviço prestado é um caso de: exportação, imunidade ou não incidência do ISSQN?*...')
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlServicoPrestado"]/div/div[2]/div/div[1]/label'))).click()

    #adicionando descrição do serviço
    print('adicionando descrição do serviço...')
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ServicoPrestado_Descricao"]'))).send_keys(dsServico)
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[7]/button/img'))).click()

    #definindo o valor da nota
    print('definindo o valor da nota...')
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Valores_ValorServico"]'))).send_keys(vlNota)

    #selecionando "Não informar nenhum valor estimado para os Tributos (Decreto 8.264/2014)" para a opção: VALOR APROXIMADO DOS TRIBUTOS
    print('selecionando "Não informar nenhum valor estimado para os Tributos (Decreto 8.264/2014)" para a opção: VALOR APROXIMADO DOS TRIBUTOS...')
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlOpcaoParaMEI"]/div/div/label'))).click()

    #finalizando Emissão
    print('finalizando Emissão e salvando os arquivos...')
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[7]/button/img'))).click()
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnProsseguir"]/img'))).click()

except Exception as e:
    screenshot_filename = "emissor.png"
    navegador.save_screenshot(screenshot_filename)
    html_filename = "emissor.html"
    with open(html_filename, "w", encoding="utf-8") as html_file:
        html_file.write(navegador.page_source)
    print(e)
    navegador.quit()
    sys.exit()

wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadXml"]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadDANFSE"]'))).click()

print('nota emitida com sucesso...')

navegador.quit()

print("finalizado com sucesso")
