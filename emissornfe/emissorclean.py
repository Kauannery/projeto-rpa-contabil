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
import sys

############################Configurações##################################
# --- DADOS REAIS DO PRESTADOR (SEUS DADOS) ---
dsEmissorCNPJ = '62018490000100' # OBRIGATGÓRIO: Use seus dados reais
dsEmissorPass = 'Teste123'       # OBRIGATÓRIO: Use seus dados reais

# --- DADOS DE TESTE (VÁLIDOS E FICTÍCIOS) ---
vlNota = '1500' # Ex: R$ 15,00
cdTomador = '06990590000123' # CNPJ da Google
dsTomadorCEP = '01311000' # CEP de teste (Av. Paulista - SP)
dsBuscaMunicipio = 'São Paulo' # Município de teste
dsMunicipio = 'São Paulo/SP' # Município de teste
dsTributario = '01.07.01' # Exemplo de "Serviços de elaboração de programas de computadores"
dsServico = 'SERVICO DE TESTE DE AUTOMACAO PARA PROJETO TCC - DESCONSIDERAR'

# --- CONFIGURAÇÕES DO SCRIPT ---
dsDownloadPath = r'C:/Faculdade/projeto-rpa/notas_teste' 
inTerminal = False # Mantenha 'False' para assistir o robô
###########################################################################

# =====================================================================================
# SEÇÃO 1: PREPARAÇÃO DO AMBIENTE (CORREÇÃO v19.A)
# OBJETIVO: Normalizar o caminho da pasta de download para evitar
#           erros de barras misturadas ( / e \ ).
# =====================================================================================
data_atual = datetime.date.today()
ano = str(data_atual.year)
mes = str(data_atual.month).zfill(2)

# --- INÍCIO DA CORREÇÃO v19.A ---
# Normaliza o caminho para o formato correto do Windows (usando '\')
dsDownloadPath_normalized = os.path.normpath(dsDownloadPath)
dsCompleteDownloadPath = os.path.join(dsDownloadPath_normalized, ano, mes)
# --- FIM DA CORREÇÃO v19.A ---

os.makedirs(dsCompleteDownloadPath, exist_ok=True)
print(f"as notas serão salvas em: {dsCompleteDownloadPath}")
# =====================================================================================
# FIM DA SEÇÃO 1
# =====================================================================================

# =====================================================================================
# SEÇÃO DE CONFIGURAÇÃO DO CHROME (CORREÇÃO v19.B)
# OBJETIVO: Forçar o Chrome a baixar os arquivos (especialmente o .xml)
#           sem bloqueios de segurança e para a pasta correta.
# =====================================================================================

# --- INÍCIO DA CORREÇÃO v19.B ---
prefs = {
    "download.default_directory": dsCompleteDownloadPath, # Caminho corrigido
    "download.prompt_for_download": False, # Não perguntar onde salvar
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False, # Desabilita o "Safe Browsing"
    "profile.default_content_settings.popups": 0,
    "profile.default_content_setting_values.automatic_downloads": 1 # Permite downloads automáticos
}

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", prefs)

# Argumentos para desabilitar popups e proteções
chrome_options.add_argument('--safebrowsing-disable-download-protection')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--ignore-certificate-errors')
# --- FIM DA CORREÇÃO v19.B ---

chrome_options.add_argument('--window-size=1920,1080')
if inTerminal:
    chrome_options.add_argument('--headless=chrome')
# =====================================================================================
# FIM DA SEÇÃO DE CONFIGURAÇÃO DO CHROME
# =====================================================================================

# Iniciando o webdriver
navegador = webdriver.Chrome(options=chrome_options)
# Aumentamos o tempo de espera padrão para 15 segundos para dar conta de lentidão
wait = WebDriverWait(navegador, 15) 

# Iniciando a emissão de nota fiscal
navegador.get('https://www.producaorestrita.nfse.gov.br/EmissorNacional/Login')

print("aceitando site inseguro...")
try:
    # Lida com a tela de "site inseguro"
    wait_seguranca = WebDriverWait(navegador, 3) # Espera curta
    wait_seguranca.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="details-button"]'))).click()
    wait_seguranca.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="proceed-link"]'))).click()
except Exception as e:
    print("Não havia página de segurança.")

try:
    # =====================================================================================
    # SEÇÃO 1: LOGIN E DATA DE COMPETÊNCIA
    # =====================================================================================
    print("logando no EmissorNacional...")
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Inscricao"]'))).send_keys(dsEmissorCNPJ)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Senha"]'))).send_keys(dsEmissorPass)
    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button'))).click()

    print("iniciando emissão de nova nota...")
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wgtAcessoRapido"]/div[2]/a[1]/img'))).click()

    print("calculando e preenchendo data de competência...")
    primeiro_dia_do_mes = data_atual.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_do_mes - datetime.timedelta(days=1)
    data_formatada = ultimo_dia_mes_anterior.strftime("%d/%m/%Y")
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="DataCompetencia"]'))).send_keys(data_formatada)
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="DataCompetencia"]'))).send_keys(Keys.TAB)

    # =====================================================================================
    # SEÇÃO 2: DADOS DO TOMADOR (CLIENTE)
    # =====================================================================================
    print("selecionando tomador baseado no CNPJ...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnMaisInfoEmitente"]')))
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="pnlTomador"]/div[1]/div/div/div[2]/label/span/i'))).click()
    
    # Bloco (v4) - Clica no botão "Usar Histórico"
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_Tomador_Inscricao_historico"]'))).click() 
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    
    # Formata o CNPJ para '06.990.590/0001-23'
    print(f"Formatando CNPJ {cdTomador} para busca no histórico...")
    cnpj_formatado = f"{cdTomador[:2]}.{cdTomador[2:5]}.{cdTomador[5:8]}/{cdTomador[8:12]}-{cdTomador[12:]}"
    print(f"Buscando pela LINHA DA TABELA que contém: {cnpj_formatado}")

    # Encontra a LINHA DA TABELA (<tr>) que contém o CNPJ e clica nela
    xpath_linha_cliente = f"//tr[contains(., '{cnpj_formatado}')]"
    print("Linha da tabela encontrada. Clicando na linha...")
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_linha_cliente))).click()

    print("Clicando em Importar...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnImportar"]'))).click()

    # =====================================================================================
    # SEÇÃO: DADOS DO TOMADOR (SUB-SEÇÃO: PREENCHIMENTO DO CEP) - CORREÇÃO v17
    # OBJETIVO: Avançar da tela de dados do tomador.
    # PROBLEMA: O botão "Avançar" (id="btnAvancar") é "não clicável" pelo Selenium
    #           (obscurecido, etc.), mesmo após o scroll.
    # SOLUÇÃO (v17): Usar um clique "forçado" via JavaScript (arguments[0].click()),
    #                que é mais robusto e bypassa as verificações de clicabilidade.
    # =====================================================================================
    # print("[DEV] Preenchendo campo de CEP de teste...")
    try:
        # 1. Encontra o CAMPO de input do CEP e digita o CEP de teste
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Tomador_EnderecoNacional_CEP"]'))).send_keys(dsTomadorCEP)
        
        # 2. Clica no BOTÃO de busca do CEP
        # print("[DEV] Clicando no botão de busca de CEP...")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_Tomador_EnderecoNacional_CEP"]'))).click()

    except Exception as e_cep:
        print(f"[DEV] Erro ao preencher ou buscar CEP: {e_cep}") # ERRO MANTIDO - IMPORTANTE

    # print("[DEV] Aguardando (se houver) modal de loading após busca de CEP...")
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    
    print("Avançando da aba 'Pessoas'...")
    xpath_avancar_pessoas = "//*[@id='btnAvancar']" # ID está correto!
    
    try:
        # Tenta o método de scroll E clique via JavaScript (Plano A)
        elemento_avancar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_pessoas)))
        # print("[DEV] Botão 'Avançar' (Pessoas) encontrado. Rolando e forçando clique via JS...")
        
        # Esta linha faz o scroll e o clique de forma forçada
        navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'}); arguments[0].click();", elemento_avancar)
        time.sleep(0.5) # Pausa para a página carregar
    
    except Exception as e_scroll:
        # Se o JS falhar, tenta o clique direto (Plano B)
        print(f"[DEV] Clique JS falhou, tentando clique direto. Erro: {e_scroll}") # ERRO MANTIDO - IMPORTANTE
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_pessoas))).click()
    
    # print("[DEV] Avançou da aba 'Pessoas' para a aba 'Serviço'.")
    # =====================================================================================
    # FIM DA SEÇÃO: DADOS DO TOMADOR
    # =====================================================================================

    # =====================================================================================
    # SEÇÃO 3: DADOS DO SERVIÇO
    # =====================================================================================
    print("selecionando Município de prestação do Serviço...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlLocalPrestacao"]/div/div/div[2]/div/span[1]/span[1]/span'))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(dsBuscaMunicipio)
    wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '"+dsMunicipio+"')]"))).click()

    print("selecionando código de tributação nacional...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlServicoPrestado"]/div/div[1]/div/div/span[1]/span[1]/span'))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(dsTributario)
    wait.until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="select2-ServicoPrestado_CodigoTributacaoNacional-results"]/li'),dsTributario))
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(Keys.TAB)

    # Bloco (v5) - Clica no "Não" do ISSQN
    print('selecionando "Não" para a opção: O serviço prestado é um caso de: exportação, imunidade ou não incidência do ISSQN?*...')
    # print("[DEV] Aguardando qualquer modal de loading desaparecer...")
    time.sleep(0.5)
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    xpath_nao = "//label[normalize-space()='Não']"
    # print("[DEV] Procurando o botão 'Não'...")
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_nao))).click()
    # print("[DEV] Botão 'Não' clicado.")

    # Bloco (v14) - Preenche a descrição e clica no PRIMEIRO "Avançar"
    print('adicionando descrição do serviço...')
    # print("[DEV] Aguardando (se houver) modal de loading após o clique no 'Não'...")
    time.sleep(0.5) 
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    
    # print("[DEV] Preenchendo a descrição...")
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ServicoPrestado_Descricao"]'))).send_keys(dsServico)
    
    # print("[DEV] Aguardando (se houver) modal de loading após preencher a descrição...")
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    
    print("Avançando da aba 'Serviço'...")
    # XPath v14 (baseado no seu HTML: <button>...<span>Avançar</span>...</button>)
    xpath_avancar_servico = "//button[span[text()='Avançar']]" 
    
    try:
        elemento_avancar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_servico)))
        # print("[DEV] Botão 'Avançar' (Serviço) encontrado. Rolando até ele...")
        navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento_avancar)
        time.sleep(0.5)
        elemento_avancar.click()
    except Exception as e_scroll:
        print(f"[DEV] Scroll falhou, tentando clique direto. Erro: {e_scroll}") # ERRO MANTIDO - IMPORTANTE
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_servico))).click()
        
    # print("[DEV] Clicou no botão 'Avançar' da aba 'Serviço'.")

    # =====================================================================================
    # SEÇÃO 4: VALORES E EMISSÃO
    # =====================================================================================
    print('definindo o valor da nota...')
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Valores_ValorServico"]'))).send_keys(vlNota)

    print('selecionando "Não informar nenhum valor estimado..." (Opção MEI)...')
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlOpcaoParaMEI"]/div/div/label'))).click()

    # Bloco (v18) - Clica no SEGUNDO "Avançar" (aba Valores) e em "Emitir"
    print('finalizando Emissão e salvando os arquivos...')
    # print("[DEV] Aguardando (se houver) modal de loading após clicar na opção MEI...")
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    
    print("Avançando da aba 'Valores'...")
    # (v18) O HTML deste botão é idêntico ao anterior.
    xpath_avancar_valores = "//button[span[text()='Avançar']]"
    
    try:
        elemento_avancar_2 = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_valores)))
        # print("[DEV] Botão 'Avançar >' encontrado. Rolando até ele...")
        navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento_avancar_2)
        time.sleep(0.5)
        elemento_avancar_2.click()
    except Exception as e_scroll_2:
        print(f"[DEV] Scroll falhou, tentando clique direto. Erro: {e_scroll_2}") # ERRO MANTIDO - IMPORTANTE
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_valores))).click()
        
    # print("[DEV] Clicou no botão 'Avançar >' da aba 'Valores'.")

    # Clica em "Emitir NFS-e"
    # print("[DEV] Aguardando modal de loading após clicar em 'Avançar >'...")
    wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
    
    print("Clicando no botão final 'Prosseguir' (Emitir NFS-e)...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnProsseguir"]/img'))).click()

# =====================================================================================
# SEÇÃO 5: DOWNLOAD E FINALIZAÇÃO
# =====================================================================================
except Exception as e:
    # Se qualquer coisa no 'try' falhar, o script vem para cá
    print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print("!!!!!!!!!!!!!! OCORREU UM ERRO !!!!!!!!!!!!!!")
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n")
    print(e)
    
    print("Salvando screenshot e HTML da página do erro...")
    screenshot_filename = "emissor_ERRO.png"
    navegador.save_screenshot(screenshot_filename)
    html_filename = "emissor_ERRO.html"
    with open(html_filename, "w", encoding="utf-8") as html_file:
        html_file.write(navegador.page_source)
    print(f"Arquivos de erro salvos como: {screenshot_filename} e {html_filename}")
    
    navegador.quit()
    sys.exit() # Para o script

# Se o 'try' foi um sucesso, o script continua aqui
print("\n====================================================")
print("NOTA EMITIDA COM SUCESSO (no ambiente de teste)!")
print("====================================================\n")

print("Baixando arquivos XML e PDF da nota...")
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadXml"]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadDANFSE"]'))).click()

# Adiciona uma pausa final para garantir que os downloads comecem
time.sleep(5) 
print('Downloads concluídos.')

navegador.quit()
print("Processo finalizado com sucesso.")
