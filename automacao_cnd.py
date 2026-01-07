# =============================================================================
# IMPORTAÇÃO DAS BIBLIOTECAS
# =============================================================================
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys # <--- IMPORTAÇÃO ADICIONADA
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =============================================================================
# CONFIGURAÇÕES DO SCRIPT
# =============================================================================
ARQUIVO_ENTRADA = "empresas.xlsx"
PASTA_RESULTADOS = "resultados_cnd_final"
URL_SERVICO = "https://servicos.receitafederal.gov.br/servico/certidoes/#/home/cnpj"

# =============================================================================
# INÍCIO DA EXECUÇÃO DO ROBÔ
# =============================================================================

if not os.path.exists(PASTA_RESULTADOS):
    os.makedirs(PASTA_RESULTADOS)
    print(f"Pasta '{PASTA_RESULTADOS}' criada com sucesso.")

print(f"Lendo o arquivo '{ARQUIVO_ENTRADA}'...")
try:
    df = pd.read_excel(ARQUIVO_ENTRADA, dtype={'CNPJ': str})
except FileNotFoundError:
    print(f"ERRO: O arquivo '{ARQUIVO_ENTRADA}' não foi encontrado.")
    exit()

print("Configurando o navegador...")
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

for index, row in df.iterrows():
    cnpj = row['CNPJ']
    cnpj_numeros = ''.join(filter(str.isdigit, cnpj))
    
    print("-" * 50)
    print(f"Processando CNPJ: {cnpj}")

    try:
        navegador.get(URL_SERVICO)
        wait = WebDriverWait(navegador, 20)

        # Lidar com o Banner de Cookies
        try:
            print("Procurando pelo banner de cookies...")
            botao_cookies = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continuar e fechar')]")))
            navegador.execute_script("arguments[0].click();", botao_cookies)
            print("Banner de cookies aceito via JS.")
            time.sleep(1)
        except Exception:
            print("Banner de cookies não encontrado ou já aceito. Seguindo em frente.")

        print("Aguardando campo CNPJ ficar visível...")
        campo_cnpj = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Informe o CNPJ']")))
        
        print("Clicando no campo via JS...")
        navegador.execute_script("arguments[0].click();", campo_cnpj)
        time.sleep(0.5)
        
        print("Preenchendo CNPJ com 'send_keys'...")
        campo_cnpj.send_keys(cnpj_numeros)
        print("CNPJ preenchido.")
        time.sleep(0.5) # Pausa curta

        # --- NOVA ETAPA: SIMULAR "TAB" ---
        print("Simulando 'Tab' para disparar a validação do campo...")
        campo_cnpj.send_keys(Keys.TAB) # Simula o usuário apertando Tab
        time.sleep(1) # Espera a validação (a mensagem de erro sumir)

        print("Clicando no botão 'Consultar Certidão' via JS...")
        botao_consultar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Consultar Certidão')]")))
        navegador.execute_script("arguments[0].click();", botao_consultar)
        
        print("Aguardando resultado...")
        time.sleep(5) # Tempo para a consulta ser processada

        nome_arquivo = f"Certidao_{cnpj_numeros}.png"
        caminho_arquivo = os.path.join(PASTA_RESULTADOS, nome_arquivo)
        navegador.save_screenshot(caminho_arquivo)
        
        print(f"Consulta finalizada. Screenshot salvo em: {caminho_arquivo}")

        df.loc[index, 'Status'] = 'Consulta Realizada com Sucesso'
        df.loc[index, 'Arquivo_Gerado'] = caminho_arquivo

    except Exception as e:
        print(f"ERRO ao processar o CNPJ {cnpj}: {e}")
        nome_arquivo_erro = f"ERRO_{cnpj_numeros}.png"
        caminho_arquivo_erro = os.path.join(PASTA_RESULTADOS, nome_arquivo_erro)
        navegador.save_screenshot(caminho_arquivo_erro)
        print(f"Screenshot do erro salvo em: {caminho_arquivo_erro}")
        df.loc[index, 'Status'] = 'Erro na Consulta'
        df.loc[index, 'Arquivo_Gerado'] = caminho_arquivo_erro

print("-" * 50)
print("Salvando os resultados finais na planilha...")
df.to_excel(ARQUIVO_ENTRADA, index=False)
navegador.quit()
print("Automação 100% concluída com sucesso!")