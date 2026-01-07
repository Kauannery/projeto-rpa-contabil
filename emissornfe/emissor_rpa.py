#!/usr/bin/env python
# NOME DESTE ARQUIVO: emissor_rpa.py

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

# =====================================================================================
# FUNÇÃO PRINCIPAL DO ROBÔ DE EMISSÃO
# Esta é a função que o nosso dashboard (Streamlit) vai chamar.
# =====================================================================================
def executar_emissao(config_rpa):
    """
    Função principal que executa o robô de emissão de NFS-e.
    Recebe um dicionário 'config_rpa' com todos os dados necessários.
    Retorna um dicionário com o resultado da operação.
    """
    
    print(">>> [ROBÔ] Iniciando processo de emissão...")
    
    # --- 1. PREPARAÇÃO DO AMBIENTE (v19.A) ---
    data_atual = datetime.date.today()
    ano = str(data_atual.year)
    mes = str(data_atual.month).zfill(2)
    
    dsDownloadPath_normalized = os.path.normpath(config_rpa["dsDownloadPath"])
    dsCompleteDownloadPath = os.path.join(dsDownloadPath_normalized, ano, mes)
    
    os.makedirs(dsCompleteDownloadPath, exist_ok=True)
    print(f"as notas serão salvas em: {dsCompleteDownloadPath}")

    # --- 2. CONFIGURAÇÃO DO CHROME (v19.B) ---
    prefs = {
        "download.default_directory": dsCompleteDownloadPath,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--safebrowsing-disable-download-protection')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--window-size=1920,1080')
    if config_rpa["inTerminal"]:
        chrome_options.add_argument('--headless=chrome')

    # Iniciando o webdriver
    navegador = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(navegador, 15) 

    # --- 3. EXECUÇÃO DO PROCESSO (try...except) ---
    try:
        navegador.get('https://www.producaorestrita.nfse.gov.br/EmissorNacional/Login')

        print("aceitando site inseguro...")
        try:
            wait_seguranca = WebDriverWait(navegador, 3)
            wait_seguranca.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="details-button"]'))).click()
            wait_seguranca.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="proceed-link"]'))).click()
        except Exception as e:
            print("Não havia página de segurança.")

        # =====================================================================================
        # SEÇÃO 1: LOGIN E DATA DE COMPETÊNCIA
        # =====================================================================================
        print("logando no EmissorNacional...")
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Inscricao"]'))).send_keys(config_rpa["dsEmissorCNPJ"])
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Senha"]'))).send_keys(config_rpa["dsEmissorPass"])
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
        # SEÇÃO 2: DADOS DO TOMADOR (CLIENTE) (v4, v17)
        # =====================================================================================
        print("selecionando tomador baseado no CNPJ...")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnMaisInfoEmitente"]')))
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="pnlTomador"]/div[1]/div/div/div[2]/label/span/i'))).click()
        
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_Tomador_Inscricao_historico"]'))).click() 
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        
        print(f"Formatando CNPJ {config_rpa['cdTomador']} para busca no histórico...")
        cnpj_formatado = f"{config_rpa['cdTomador'][:2]}.{config_rpa['cdTomador'][2:5]}.{config_rpa['cdTomador'][5:8]}/{config_rpa['cdTomador'][8:12]}-{config_rpa['cdTomador'][12:]}"
        
        xpath_linha_cliente = f"//tr[contains(., '{cnpj_formatado}')]"
        print("Linha da tabela encontrada. Clicando na linha...")
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_linha_cliente))).click()

        print("Clicando em Importar...")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnImportar"]'))).click()

        # (v17) Preenche o CEP e clica em "Avançar" com JS forçado
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Tomador_EnderecoNacional_CEP"]'))).send_keys(config_rpa["dsTomadorCEP"])
            wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_Tomador_EnderecoNacional_CEP"]'))).click()
        except Exception as e_cep:
            print(f"Erro ao preencher ou buscar CEP: {e_cep}") # ERRO MANTIDO - IMPORTANTE

        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        
        print("Avançando da aba 'Pessoas'...")
        xpath_avancar_pessoas = "//*[@id='btnAvancar']"
        
        try:
            elemento_avancar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_pessoas)))
            navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'}); arguments[0].click();", elemento_avancar)
            time.sleep(0.5)
        except Exception as e_scroll:
            print(f"Clique JS falhou, tentando clique direto. Erro: {e_scroll}") # ERRO MANTIDO - IMPORTANTE
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_pessoas))).click()

        # =====================================================================================
        # SEÇÃO 3: DADOS DO SERVIÇO (v5, v14)
        # =====================================================================================
        print("selecionando Município de prestação do Serviço...")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlLocalPrestacao"]/div/div/div[2]/div/span[1]/span[1]/span'))).click()
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(config_rpa["dsBuscaMunicipio"])
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '"+config_rpa["dsMunicipio"]+"')]"))).click()

        print("selecionando código de tributação nacional...")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlServicoPrestado"]/div/div[1]/div/div/span[1]/span[1]/span'))).click()
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(config_rpa["dsTributario"])
        wait.until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="select2-ServicoPrestado_CodigoTributacaoNacional-results"]/li'),config_rpa["dsTributario"]))
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))).send_keys(Keys.TAB)

        print('selecionando "Não" para a opção ISSQN...')
        time.sleep(0.5)
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        xpath_nao = "//label[normalize-space()='Não']"
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_nao))).click()

        print('adicionando descrição do serviço...')
        time.sleep(0.5) 
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ServicoPrestado_Descricao"]'))).send_keys(config_rpa["dsServico"])
        
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        
        print("Avançando da aba 'Serviço'...")
        xpath_avancar_servico = "//button[span[text()='Avançar']]" 
        
        try:
            elemento_avancar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_servico)))
            navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento_avancar)
            time.sleep(0.5)
            elemento_avancar.click()
        except Exception as e_scroll:
            print(f"Scroll falhou, tentando clique direto. Erro: {e_scroll}")
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_servico))).click()
            
        # =====================================================================================
        # SEÇÃO 4: VALORES E EMISSÃO (v18)
        # =====================================================================================
        print('definindo o valor da nota...')
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Valores_ValorServico"]'))).send_keys(config_rpa["vlNota"])

        print('selecionando "Não informar nenhum valor estimado..." (Opção MEI)...')
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlOpcaoParaMEI"]/div/div/label'))).click()

        print('finalizando Emissão e salvando os arquivos...')
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        
        print("Avançando da aba 'Valores'...")
        xpath_avancar_valores = "//button[span[text()='Avançar']]" # (v18)
        
        try:
            elemento_avancar_2 = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_valores)))
            navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento_avancar_2)
            time.sleep(0.5)
            elemento_avancar_2.click()
        except Exception as e_scroll_2:
            print(f"Scroll falhou, tentando clique direto. Erro: {e_scroll_2}")
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_valores))).click()
            
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        
        print("Clicando no botão final 'Prosseguir' (Emitir NFS-e)...")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnProsseguir"]/img'))).click()

    # =====================================================================================
    # SEÇÃO 5: DOWNLOAD E FINALIZAÇÃO
    # =====================================================================================
    except Exception as e:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO !!!!!!!!!!!!!!\n{e}")
        
        print("Salvando screenshot e HTML da página do erro...")
        screenshot_filename = "emissor_ERRO.png"
        navegador.save_screenshot(screenshot_filename)
        html_filename = "emissor_ERRO.html"
        with open(html_filename, "w", encoding="utf-8") as html_file:
            html_file.write(navegador.page_source)
        
        navegador.quit()
        # Retorna um resultado de erro
        return {"sucesso": False, "mensagem": str(e), "screenshot": screenshot_filename}

    # Se o 'try' foi um sucesso, o script continua aqui
    print("\n====================================================")
    print("NOTA EMITIDA COM SUCESSO (no ambiente de teste)!")
    print("====================================================\n")

    print("Baixando arquivos XML e PDF da nota...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadXml"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadDANFSE"]'))).click()

    time.sleep(5) 
    print('Downloads concluídos.')
    navegador.quit()
    print("Processo finalizado com sucesso.")
    
    # Retorna um resultado de sucesso
    return {"sucesso": True, "mensagem": f"Nota fiscal emitida com sucesso! Os arquivos foram salvos em: {dsCompleteDownloadPath}"}