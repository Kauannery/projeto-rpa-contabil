#!/usr/bin/env python
# NOME DESTE ARQUIVO: app_tcc.py
# Este arquivo contém AMBOS o robô (RPA) e a interface (Streamlit).

# --- Importações para o Robô (Selenium) ---
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

# --- Importações para o Dashboard (Streamlit) ---
import streamlit as st

# --- (v27) Importação para Análise de PDF ---
import fitz  # PyMuPDF

# =====================================================================================
# INSTRUÇÕES DO PROJETO (README)
# (Ocultado para brevidade)
# =====================================================================================


# =====================================================================================
# FUNÇÃO 1: ROBÔ DE EMISSÃO DE NFS-e (O "MOTOR" 1)
# (Esta função está correta e não foi alterada)
# =====================================================================================
def executar_emissao(config_rpa):
    """
    Função principal que executa o robô de emissão de NFS-e.
    """
    
    print(">>> [ROBÔ NFS-e] Iniciando...")
    
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
        # (Ocultado para brevidade, está correto)
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
        # (Ocultado para brevidade, está correto)
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
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Tomador_EnderecoNacional_CEP"]'))).send_keys(config_rpa["dsTomadorCEP"])
            wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_Tomador_EnderecoNacional_CEP"]'))).click()
        except Exception as e_cep:
            print(f"[DEV] Erro ao preencher ou buscar CEP: {e_cep}")
        wait.until_not(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalLoading"]')))
        print("Avançando da aba 'Pessoas'...")
        xpath_avancar_pessoas = "//*[@id='btnAvancar']"
        try:
            elemento_avancar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_avancar_pessoas)))
            navegador.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'}); arguments[0].click();", elemento_avancar)
            time.sleep(0.5)
        except Exception as e_scroll:
            print(f"Clique JS falhou, tentando clique direto. Erro: {e_scroll}")
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_avancar_pessoas))).click()

        # =====================================================================================
        # SEÇÃO 3: DADOS DO SERVIÇO (v5, v14)
        # (Ocultado para brevidade, está correto)
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
        # (Ocultado para brevidade, está correto)
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
    # SEÇÃO 5: DOWNLOAD E FINALIZAÇÃO (NFS-e)
    # (Ocultado para brevidade, está correto)
    # =====================================================================================
    except Exception as e:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO (NFS-e) !!!!!!!!!!!!!!\n{e}")
        print("Salvando screenshot e HTML da página do erro...")
        dir_atual = os.path.dirname(os.path.abspath(__file__))
        screenshot_filename = os.path.join(dir_atual, "emissor_ERRO.png")
        html_filename = os.path.join(dir_atual, "emissor_ERRO.html")
        navegador.save_screenshot(screenshot_filename)
        with open(html_filename, "w", encoding="utf-8") as html_file:
            html_file.write(navegador.page_source)
        navegador.quit()
        return {"sucesso": False, "mensagem": str(e), "screenshot": screenshot_filename}

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
    return {"sucesso": True, "mensagem": f"Nota fiscal emitida com sucesso! Os arquivos foram salvos em: {dsCompleteDownloadPath}"}


# =====================================================================================
# FUNÇÃO 2: ROBÔ DE CONSULTA DE CND TRABALHISTA (O "MOTOR" 2) - v27 (Atualizado)
# =====================================================================================
def executar_consulta_cnd_tst(cnpj, download_path):
    """
    Executa a consulta de CNDT (Trabalhista) no portal do TST.
    (v27): Atualizado com "File Watcher" para renomear o PDF baixado.
    """
    print(">>> [ROBÔ CND-TST] Iniciando consulta...")
    dsDownloadPath_normalized = os.path.normpath(download_path)
    os.makedirs(dsDownloadPath_normalized, exist_ok=True)
    print(f"A certidão será salva em: {dsDownloadPath_normalized}")

    prefs = {
        "download.default_directory": dsDownloadPath_normalized,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": False
    }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--window-size=1024,768')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--ignore-certificate-errors')

    navegador = webdriver.Chrome(options=chrome_options)
    
    try:
        navegador.get('https://www.tst.jus.br/certidao')
        wait_longo = WebDriverWait(navegador, 180) 
        wait_curto = WebDriverWait(navegador, 5)

        try:
            print("Procurando o banner de cookies...")
            xpath_aceitar_cookie = "//*[@id='cookie-save-button-lgpd']"
            botao_aceitar = wait_curto.until(EC.element_to_be_clickable((By.XPATH, xpath_aceitar_cookie)))
            botao_aceitar.click()
            print("Banner de cookies aceito.")
            time.sleep(1) 
        except Exception as e_cookie:
            print("Banner de cookies não encontrado ou já foi aceito.")

        print("Procurando o formulário e preenchendo o CNPJ...")
        wait_longo.until(EC.presence_of_element_located((By.ID, 'corpo:formulario:cnpj'))).send_keys(cnpj)
        
        print(">>> AÇÃO DO USUÁRIO NECESSÁRIA <<<")
        print("Por favor, vá para a janela do robô (Chrome).")
        print("1. Marque a caixa 'Não sou um robô'.")
        print("2. Clique no botão 'Pesquisar'.")
        print("O robô está esperando o resultado...")
        
        xpath_botao_emitir = "//input[@value='Emitir Certidão']"
        wait_longo.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_emitir)))
        
        print("Usuário resolveu o CAPTCHA. Robô assumindo...")

        # (v27) Lógica de "vigia" de pasta
        arquivos_antes = set(os.listdir(dsDownloadPath_normalized))
        
        navegador.find_element(By.XPATH, xpath_botao_emitir).click()
        
        print("Aguardando o download do arquivo PDF...")
        novo_arquivo_path = aguardar_novo_arquivo(dsDownloadPath_normalized, arquivos_antes, 20)

        if not novo_arquivo_path:
            raise Exception("Download não concluído no tempo limite.")

        # (v27) Renomeia o arquivo
        print(f"Arquivo detectado. Renomeando...")
        cnpj_limpo = ''.join(filter(str.isdigit, cnpj))
        nome_novo_pdf = f"CND_TST_{cnpj_limpo}.pdf"
        nome_novo_completo = os.path.join(dsDownloadPath_normalized, nome_novo_pdf)
        
        time.sleep(1) # Espera o arquivo ser "solto"
        os.rename(novo_arquivo_path, nome_novo_completo)
        print(f"Arquivo renomeado para: {nome_novo_pdf}")

    except Exception as e:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO (CND-TST) !!!!!!!!!!!!!!\n{e}")
        dir_atual = os.path.dirname(os.path.abspath(__file__))
        screenshot_filename = os.path.join(dir_atual, "cnd_tst_ERRO.png")
        navegador.save_screenshot(screenshot_filename)
        navegador.quit()
        return {"sucesso": False, "mensagem": str(e), "screenshot": screenshot_filename}
    
    navegador.quit()
    print("Processo CND-TST finalizado com sucesso.")
    return {"sucesso": True, "caminho_arquivo": nome_novo_completo}

# =====================================================================================
# FUNÇÃO 3: ROBÔ DE CONSULTA DE CND SEFAZ-BA (O "MOTOR" 3) - v27 (LOTE ROBUSTO)
# =====================================================================================
def executar_consulta_cnd_sefaz_ba(lista_de_cnpjs, download_path):
    """
    (v27) Executa a consulta de CND Estadual em LOTE no portal da SEFAZ-BA.
    Usa um "vigia" de pasta para detectar downloads e renomeia o arquivo.
    """
    print(">>> [ROBÔ CND-SEFAZ] Iniciando consulta em LOTE...")

    # --- 1. PREPARAÇÃO DO AMBIENTE ---
    dsDownloadPath_normalized = os.path.normpath(download_path)
    os.makedirs(dsDownloadPath_normalized, exist_ok=True)
    print(f"A certidão será salva em: {dsDownloadPath_normalized}")

    # --- 2. CONFIGURAÇÃO DO CHROME ---
    prefs = {
        "download.default_directory": dsDownloadPath_normalized,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True, 
        "safebrowsing.enabled": False
    }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--window-size=1024,768')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--ignore-certificate-errors')

    navegador = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(navegador, 15) # Espera de 15 seg por consulta
    
    # --- 3. EXECUÇÃO DO PROCESSO EM LOTE ---
    logs_sucesso = []
    logs_falha = []
    url_sefaz = 'https://servicos.sefaz.ba.gov.br/sistemas/DSCRE/Modulos/Publico/EmissaoCertidao.aspx'
    
    try:
        print(f"Iniciando loop para {len(lista_de_cnpjs)} CNPJs...")
        
        for cnpj in lista_de_cnpjs:
            try:
                print(f"--- Consultando CNPJ: {cnpj} ---")
                navegador.get(url_sefaz)
                
                # (v27) Pega a lista de arquivos ANTES do download
                arquivos_antes = set(os.listdir(dsDownloadPath_normalized))
                
                # (v22) Preenche o CNPJ
                id_campo_cnpj = "PHConteudo_TxtNumCNPJ"
                campo_cnpj = wait.until(EC.visibility_of_element_located((By.ID, id_campo_cnpj)))
                campo_cnpj.clear()
                campo_cnpj.send_keys(cnpj)
                
                # (v22) Clica em "Imprimir"
                id_botao_imprimir = "PHConteudo_btnImprimir"
                elemento_imprimir = wait.until(EC.element_to_be_clickable((By.ID, id_botao_imprimir)))
                navegador.execute_script("arguments[0].click();", elemento_imprimir) # Usando clique JS

                # (v27) Lógica de "vigia" de pasta
                print("Aguardando o download do arquivo PDF...")
                novo_arquivo_path = aguardar_novo_arquivo(dsDownloadPath_normalized, arquivos_antes, 20)
                
                if not novo_arquivo_path:
                    # Se o loop terminar sem um arquivo, verifica a página por erros
                    if "Nenhuma certidão" in navegador.page_source or "não possui" in navegador.page_source or "Inscrição Estadual inválida" in navegador.page_source:
                         print(f"Não foi possível emitir para {cnpj} (Pendências ou Inscrição inválida).")
                         raise Exception("Não foi possível emitir (Possui pendências ou Inscrição Estadual inválida na BA)")
                    else:
                        raise Exception(f"Download não concluído em 20 segundos.")

                # (v27) Renomeia o arquivo
                print(f"Arquivo detectado. Renomeando...")
                cnpj_limpo = ''.join(filter(str.isdigit, cnpj))
                nome_novo_pdf = f"CND_SEFAZ_BA_{cnpj_limpo}.pdf"
                nome_novo_completo = os.path.join(dsDownloadPath_normalized, nome_novo_pdf)
                
                time.sleep(1) # Espera o arquivo ser "solto"
                os.rename(novo_arquivo_path, nome_novo_completo)
                
                print(f"Arquivo renomeado para: {nome_novo_pdf}")
                logs_sucesso.append({"cnpj": cnpj, "arquivo": nome_novo_completo})
            
            except Exception as e_cnpj:
                print(f"Falha ao processar CNPJ {cnpj}: {e_cnpj}")
                logs_falha.append(f"CNPJ {cnpj}: Falha no robô. ({e_cnpj})")
                continue # Pula para o próximo CNPJ

    except Exception as e_geral:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO GERAL (CND-SEFAZ) !!!!!!!!!!!!!!\n{e_geral}")
        dir_atual = os.path.dirname(os.path.abspath(__file__))
        screenshot_filename = os.path.join(dir_atual, "cnd_sefaz_ERRO.png")
        navegador.save_screenshot(screenshot_filename)
        return {"sucesso": False, "mensagem": str(e_geral), "screenshot": screenshot_filename, "sucessos": logs_sucesso, "falhas": logs_falha}
    
    finally:
        navegador.quit()
        print("Processo CND-SEFAZ em lote finalizado.")
    
    return {"sucesso": True, "mensagem": f"Consulta em lote finalizada.", "sucessos": logs_sucesso, "falhas": logs_falha}


# =====================================================================================
# FUNÇÃO 4: ROBÔ DE CONSULTA DE CND FEDERAL (O "MOTOR" 4) - v27 (Atualizado)
# =====================================================================================
def executar_consulta_cnd_federal(cnpj, download_path):
    """
    (v27) Executa a consulta de CND Federal (Receita Federal / PGFN).
    Automação "assistida" que exige que o usuário resolva o CAPTCHA.
    Atualizado com "File Watcher" para renomear o PDF baixado.
    """
    print(">>> [ROBÔ CND-FEDERAL] Iniciando consulta...")

    # --- 1. PREPARAÇÃO DO AMBIENTE ---
    dsDownloadPath_normalized = os.path.normpath(download_path)
    os.makedirs(dsDownloadPath_normalized, exist_ok=True)
    print(f"A certidão será salva em: {dsDownloadPath_normalized}")

    # --- 2. CONFIGURAÇÃO DO CHROME ---
    prefs = {
        "download.default_directory": dsDownloadPath_normalized,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": False
    }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--window-size=1024,768')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--ignore-certificate-errors')

    navegador = webdriver.Chrome(options=chrome_options)
    
    # --- 3. EXECUÇÃO DO PROCESSO (try...except) ---
    try:
        # (v24) URL CORRIGIDA
        navegador.get('https://servicos.receita.fazenda.gov.br/Servicos/certidao/PJ/Emitir')
        
        wait_longo = WebDriverWait(navegador, 180) # 3 min para o usuário
        
        # 1. Robô preenche o CNPJ
        print("Preenchendo CNPJ...")
        id_campo_cnpj = "NI" # ID CORRETO
        wait_longo.until(EC.visibility_of_element_located((By.ID, id_campo_cnpj))).send_keys(cnpj)
        
        # 2. Robô avisa e PAUSA, esperando o usuário agir
        print(">>> AÇÃO DO USUÁRIO NECESSÁRIA <<<")
        print("Por favor, vá para a janela do robô (Chrome).")
        print("1. Digite as letras do CAPTCHA.")
        print("2. Clique no botão 'Consultar'.")
        print("O robô está esperando o resultado...")
        
        # 3. Robô espera o resultado (a próxima página)
        id_link_download = "btnEmitirCertidao" # ID CORRETO
        wait_longo.until(EC.element_to_be_clickable((By.ID, id_link_download)))
        
        print("Usuário resolveu o CAPTCHA. Robô assumindo...")

        # 4. (v27) Lógica de "vigia" de pasta
        arquivos_antes = set(os.listdir(dsDownloadPath_normalized))

        elemento_imprimir = navegador.find_element(By.ID, id_link_download)
        navegador.execute_script("arguments[0].click();", elemento_imprimir) # Usando clique JS
        
        print("Aguardando o download do arquivo PDF...")
        novo_arquivo_path = aguardar_novo_arquivo(dsDownloadPath_normalized, arquivos_antes, 20)

        if not novo_arquivo_path:
            raise Exception("Download não concluído no tempo limite.")

        # 5. (v27) Renomeia o arquivo
        print(f"Arquivo detectado. Renomeando...")
        cnpj_limpo = ''.join(filter(str.isdigit, cnpj))
        nome_novo_pdf = f"CND_Federal_{cnpj_limpo}.pdf"
        nome_novo_completo = os.path.join(dsDownloadPath_normalized, nome_novo_pdf)
        
        time.sleep(1) # Espera o arquivo ser "solto"
        os.rename(novo_arquivo_path, nome_novo_completo)
        print(f"Arquivo renomeado para: {nome_novo_pdf}")

    except Exception as e:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO (CND-FEDERAL) !!!!!!!!!!!!!!\n{e}")
        dir_atual = os.path.dirname(os.path.abspath(__file__))
        screenshot_filename = os.path.join(dir_atual, "cnd_federal_ERRO.png")
        navegador.save_screenshot(screenshot_filename)
        navegador.quit()
        return {"sucesso": False, "mensagem": str(e), "screenshot": screenshot_filename}
    
    navegador.quit()
    print("Processo CND-FEDERAL finalizado com sucesso.")
    return {"sucesso": True, "caminho_arquivo": nome_novo_completo}

# =====================================================================================
# FUNÇÃO 5: ANALISADOR DE PDF DE CND (O "MÓDULO DE ANÁLISE") - v27 (NOVO)
# =====================================================================================
def analisar_pdf_cnd(caminho_pdf):
    """
    (v27) Abre um arquivo PDF de CND, extrai seu texto e verifica se
    existem pendências com base em palavras-chave.
    """
    print(f">>> [ANALISADOR PDF] Lendo o arquivo: {caminho_pdf}")
    
    # Palavras-chave que indicam uma certidão "limpa" (Negativa)
    keywords_negativa = [
        "nada consta",
        "certidão negativa",
        "negativa de débitos",
        "situação regular",
        "positiva com os mesmos efeitos de negativa", # Importante!
        "positiva com efeito de negativa" # Variação
    ]
    
    # Palavras-chave que indicam uma certidão "suja" (Positiva)
    keywords_positiva = [
        "constam débitos",
        "positiva de débitos",
        "situação irregular",
        "não foi possível emitir"
    ]
    
    try:
        # 1. Abre o PDF com o PyMuPDF (fitz)
        with fitz.open(caminho_pdf) as doc:
            texto_completo = ""
            # 2. Extrai o texto de todas as páginas
            for page in doc:
                texto_completo += page.get_text()
        
        # 3. Limpa o texto para facilitar a busca
        # (remove quebras de linha, bota em minúsculo, remove acentos)
        texto_limpo = texto_completo.lower().replace("\n", " ")
        texto_limpo = ''.join(c for c in texto_limpo if c.isalnum() or c.isspace())
        
        # 4. Procura pelas palavras-chave
        for key in keywords_negativa:
            if key in texto_limpo:
                print(f"Análise: Encontrada keyword de SUCESSO: '{key}'")
                return {"status": "Negativa", "mensagem": f"✅ NADA CONSTA (ou Positiva com Efeito de Negativa)."}
                
        for key in keywords_positiva:
            if key in texto_limpo:
                print(f"Análise: Encontrada keyword de ALERTA: '{key}'")
                return {"status": "Positiva", "mensagem": f"⚠️ POSSUI PENDÊNCIAS. Verifique o PDF."}

        # 5. Se não achar nada
        print("Análise: Nenhuma palavra-chave encontrada, retornando status desconhecido.")
        return {"status": "Desconhecido", "mensagem": f"❓ Não foi possível determinar o status. Verifique o PDF."}

    except Exception as e_pdf:
        print(f"Erro ao ler o PDF: {e_pdf}")
        return {"status": "Erro", "mensagem": f"❌ Erro ao ler o arquivo PDF: {e_pdf}"}

# =====================================================================================
# FUNÇÃO AUXILIAR: VIGIA DE PASTA (v27)
# =====================================================================================
def aguardar_novo_arquivo(pasta, arquivos_antes, tempo_limite=20):
    """
    Monitora uma pasta por um 'tempo_limite' e retorna o caminho
    completo do primeiro novo arquivo .pdf encontrado.
    """
    start_time = time.time()
    while (time.time() - start_time) < tempo_limite:
        time.sleep(0.5) # Verifica a cada meio segundo
        arquivos_depois = set(os.listdir(pasta))
        novos_arquivos = arquivos_depois - arquivos_antes
        
        if novos_arquivos:
            for nome_arquivo in novos_arquivos:
                # Ignora arquivos temporários de download do Chrome
                if not nome_arquivo.endswith(".crdownload") and nome_arquivo.endswith(".pdf"):
                    print(f"Novo arquivo detectado: {nome_arquivo}")
                    return os.path.join(pasta, nome_arquivo)
    
    # Se o tempo esgotar
    return None

# =====================================================================================
# INTERFACE DO DASHBOARD (Streamlit)
# O código abaixo é o "site" que chama as funções do robô.
# =====================================================================================

def main():
    # --- Configuração da Página ---
    st.set_page_config(
        page_title="RPA Contábil - TCC",
        page_icon="🤖",
        layout="centered",
        initial_sidebar_state="auto"
    )

    # --- Título e Descrição ---
    st.title("🤖 Protótipo RPA Contábil CONTASBOT- TCC")
    st.markdown("Bem-vindo ao dashboard de automação. Este protótipo utiliza RPA (Robotic Process Automation) para executar tarefas contábeis.")
    
    # --- Abas para os Módulos ---
    tab1, tab2 = st.tabs(["Módulo 1: Emissão de NFS-e", "Módulo 2: Consulta de CNDs"])

    # =================================================
    # ABA 1: EMISSOR DE NFS-e
    # =================================================
    with tab1:
        st.header("Módulo 1: Emissão de NFS-e (MEI)")
        st.info("Preencha os dados abaixo e clique em 'Emitir Nota' para iniciar o robô. A emissão será feita no **Ambiente de Teste** do governo (Produção Restrita).")

        # --- Formulário de Entrada de Dados ---
        with st.form(key="emissao_form"):
            st.subheader("1. Dados do Prestador (Você - MEI)")
            col1, col2 = st.columns(2)
            with col1:
                dsEmissorCNPJ = st.text_input("Seu CNPJ (Prestador)", value="62018490000100")
            with col2:
                dsEmissorPass = st.text_input("Sua Senha (do Portal)", value="Teste123", type="password")

            st.subheader("2. Dados do Tomador (Cliente)")
            col3, col4 = st.columns(2)
            with col3:
                cdTomador = st.text_input("CNPJ do Cliente (Tomador)", value="06990590000123")
            with col4:
                dsTomadorCEP = st.text_input("CEP do Cliente", value="01311000")

            st.subheader("3. Dados do Serviço")
            col5, col6 = st.columns(2)
            with col5:
                vlNota = st.text_input("Valor do Serviço (Ex: 15.00)", value="15.00")
            with col6:
                dsTributario = st.text_input("Cód. Tributação (Ex: 01.07.01)", value="01.07.01")
            
            dsServico = st.text_area(
                "Descrição do Serviço", 
                value="SERVICO DE TESTE DE AUTOMACAO PARA PROJETO TCC - DESCONSIDERAR"
            )

            st.subheader("4. Configurações do Robô (NFS-e)")
            default_path = os.path.normpath(r'C:/Faculdade/projeto-rpa/notas_teste')
            dsDownloadPath_nfse = st.text_input("Pasta para Salvar as Notas", value=default_path, key="path_nfse")
            
            inTerminal_nfse = st.checkbox("Rodar robô em segundo plano (headless)", value=False, key="headless_nfse")
            
            submit_button_nfse = st.form_submit_button(label="🚀 Emitir Nota Fiscal (Teste)")

        # --- Lógica de Execução (NFS-e) ---
        if submit_button_nfse:
            # (Código de execução do robô 1)
            st.markdown("---")
            st.info("O robô de NFS-e foi iniciado! Por favor, aguarde...")
            st.warning("Uma janela do navegador será aberta. Não mexa no mouse ou teclado até o processo terminar.")
            config_rpa = {
                "dsEmissorCNPJ": dsEmissorCNPJ, "dsEmissorPass": dsEmissorPass,
                "vlNota": vlNota.replace(',', '.'), "cdTomador": cdTomador,
                "dsTomadorCEP": dsTomadorCEP, "dsBuscaMunicipio": "São Paulo",
                "dsMunicipio": "São Paulo/SP", "dsTributario": dsTributario,
                "dsServico": dsServico, "dsDownloadPath": dsDownloadPath_nfse,
                "inTerminal": inTerminal_nfse
            }
            try:
                with st.spinner('O robô de NFS-e está trabalhando... (Isso pode levar 1-2 minutos)'):
                    resultado = executar_emissao(config_rpa)
                if resultado["sucesso"]:
                    st.success("Robô de NFS-e finalizado com sucesso!")
                    st.balloons()
                    st.json(resultado)
                else:
                    st.error(f"O robô de NFS-e falhou: {resultado['mensagem']}")
                    if "screenshot" in resultado:
                        st.image("emissor_ERRO.png", caption=f"Screenshot do erro: {resultado['screenshot']}")
            except Exception as e:
                st.error(f"Ocorreu um erro crítico ao tentar rodar o robô de NFS-e: {e}")

    # =================================================
    # ABA 2: CONSULTA DE CNDs (ATUALIZADA)
    # =================================================
    with tab2:
        st.header("Módulo 2: Consulta de Certidões (CND)")
        st.info("Este módulo automatiza a consulta de CNDs. Alguns portais exigem ajuda manual com o CAPTCHA.")

        with st.form(key="cnd_form"):
            st.subheader("1. Dados da Consulta")
            
            # (v25) Campo de CNPJ agora é uma área de texto
            st.write("CNPJs para consulta (um por linha):")
            cnpjs_para_consulta = st.text_area("CNPJs", value="06990590000123\n05429636000177", height=150)
            
            st.subheader("2. Configurações do Robô (CND)")
            default_path_cnd = os.path.normpath(r'C:/Faculdade/projeto-rpa/certidoes')
            dsDownloadPath_cnd = st.text_input("Pasta para Salvar as Certidões", value=default_path_cnd, key="path_cnd")
            
            st.markdown("---")
            # --- (v23) Layout de 3 colunas para os botões ---
            col_cnd1, col_cnd2, col_cnd3 = st.columns(3)
            with col_cnd1:
                submit_button_cnd_federal = st.form_submit_button(label="🚀 CND Federal (RFB)")
            with col_cnd2:
                # (v25) O botão SEFAZ agora é em LOTE
                submit_button_cnd_sefaz = st.form_submit_button(label="🚀 CND SEFAZ-BA (Em Lote)")
            with col_cnd3:
                submit_button_cnd_tst = st.form_submit_button(label="🚀 CND Trabalhista (TST)")
        
        # --- (v23) Outros botões ficam fora do formulário ---
        st.button("Consultar CND FGTS (Caixa) (Em Breve)", disabled=True) 

        # --- Lógica de Execução (FEDERAL) ---
        if submit_button_cnd_federal:
            st.markdown("---")
            st.info("O robô da CND Federal foi iniciado!")
            st.warning("⚠️ **AÇÃO NECESSÁRIA:** Uma janela do robô será aberta. Por favor, **digite as letras do CAPTCHA** e clique em 'Consultar'. O robô está pausado e esperando por você.")

            cnpj_federal = cnpjs_para_consulta.split('\n')[0].strip()
            if not cnpj_federal:
                st.error("Por favor, insira pelo menos um CNPJ.")
            else:
                try:
                    with st.spinner('O robô está aguardando você resolver o CAPTCHA... (Limite de 3 min)'):
                        resultado_cnd = executar_consulta_cnd_federal(cnpj_federal, dsDownloadPath_cnd)
                    
                    if resultado_cnd["sucesso"]:
                        st.success("Robô da CND Federal finalizado com sucesso!")
                        st.balloons()
                        
                        # (v27) Chama a função de análise
                        st.subheader("Análise Automática do PDF:")
                        resultado_analise = analisar_pdf_cnd(resultado_cnd["caminho_arquivo"])
                        if resultado_analise["status"] == "Negativa":
                            st.success(resultado_analise["mensagem"])
                        elif resultado_analise["status"] == "Positiva":
                            st.warning(resultado_analise["mensagem"])
                        else:
                            st.info(resultado_analise["mensagem"])
                    
                    else:
                        st.error(f"O robô da CND Federal falhou: {resultado_cnd['mensagem']}")
                        if "screenshot" in resultado_cnd:
                            st.image("cnd_federal_ERRO.png", caption=f"Screenshot do erro: {resultado_cnd['screenshot']}")
                
                except Exception as e:
                    st.error(f"Ocorreu um erro crítico ao tentar rodar o robô da CND Federal: {e}")

        # --- Lógica de Execução (SEFAZ-BA) ---
        if submit_button_cnd_sefaz:
            st.markdown("---")
            st.info("O robô da SEFAZ-BA (em lote) foi iniciado! Este robô é 100% automático.")

            lista_de_cnpjs = [cnpj.strip() for cnpj in cnpjs_para_consulta.split('\n') if cnpj.strip()]
            
            if not lista_de_cnpjs:
                st.error("Por favor, insira pelo menos um CNPJ na área de texto.")
            else:
                try:
                    with st.spinner(f'O robô da SEFAZ-BA está trabalhando em {len(lista_de_cnpjs)} CNPJ(s)...'):
                        resultado_cnd = executar_consulta_cnd_sefaz_ba(lista_de_cnpjs, dsDownloadPath_cnd)
                    
                    st.success(f"Consulta em lote da SEFAZ-BA finalizada!")
                    
                    # (v27) Mostra um resumo claro com ANÁLISE DE PDF
                    with st.expander("Ver Resumo Detalhado da Execução em Lote"):
                        st.subheader(f"Sucessos ({len(resultado_cnd['sucessos'])}):")
                        for log in resultado_cnd['sucessos']:
                            # Chama a análise para cada arquivo
                            analise = analisar_pdf_cnd(log["arquivo"])
                            if analise["status"] == "Negativa":
                                st.write(f"✅ **{log['cnpj']}:** {analise['mensagem']}")
                            else:
                                st.write(f"⚠️ **{log['cnpj']}:** {analise['mensagem']}")
                        
                        st.subheader(f"Falhas ({len(resultado_cnd['falhas'])}):")
                        for log in resultado_cnd['falhas']:
                            st.write(f"❌ {log}")

                except Exception as e:
                    st.error(f"Ocorreu um erro crítico ao tentar rodar o robô da SEFAZ-BA: {e}")

        # --- Lógica de Execução (TST) ---
        if submit_button_cnd_tst:
            st.markdown("---")
            st.info("O robô do TST foi iniciado!")
            st.warning("⚠️ **AÇÃO NECESSÁRIA:** Uma janela do robô será aberta. Por favor, **resolva o CAPTCHA ('Não sou um robô')** e clique em 'Pesquisar'. O robô está pausado e esperando por você.")

            cnpj_tst = cnpjs_para_consulta.split('\n')[0].strip()
            
            if not cnpj_tst:
                st.error("Por favor, insira pelo menos um CNPJ.")
            else:
                try:
                    with st.spinner('O robô do TST está aguardando você resolver o CAPTCHA... (Limite de 3 min)'):
                        resultado_cnd = executar_consulta_cnd_tst(cnpj_tst, dsDownloadPath_cnd)
                    
                    if resultado_cnd["sucesso"]:
                        st.success("Robô do TST finalizado com sucesso!")
                        st.balloons()
                        
                        # (v27) Chama a função de análise
                        st.subheader("Análise Automática do PDF:")
                        resultado_analise = analisar_pdf_cnd(resultado_cnd["caminho_arquivo"])
                        if resultado_analise["status"] == "Negativa":
                            st.success(resultado_analise["mensagem"])
                        elif resultado_analise["status"] == "Positiva":
                            st.warning(resultado_analise["mensagem"])
                        else:
                            st.info(resultado_analise["mensagem"])
                    else:
                        st.error(f"O robô do TST falhou: {resultado_cnd['mensagem']}")
                        if "screenshot" in resultado_cnd:
                            st.image("cnd_tst_ERRO.png", caption=f"Screenshot do erro: {resultado_cnd['screenshot']}")
            
                except Exception as e:
                    st.error(f"Ocorreu um erro crítico ao tentar rodar o robô do TST: {e}")


# Esta linha é o "ponto de entrada" que executa o dashboard quando você roda o script
if __name__ == "__main__":
    main()