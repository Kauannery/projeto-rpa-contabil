#!/usr/bin/env python
# NOME DESTE ARQUIVO: app_modulo3_prefeitura.py
# (v41) - Corrigido o ID do sub-menu "NFS-e" (id="recurso-issqn-nfse")

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os
import time
import sys
import datetime # Importado para as datas

# --- Importações para o Dashboard (Streamlit) ---
import streamlit as st

# =====================================================================================
# CONFIGURAÇÕES DO MÓDULO 3
# =====================================================================================
# (PREENCHA COM SEU LOGIN E SENHA REAIS DO PORTAL DE PRODUÇÃO)
PREFEITURA_LOGIN_CNPJ = "SEU_CPF_OU_CNPJ_REAL_AQUI" # Seu login real
PREFEITURA_SENHA = "SUA_SENHA_REAL_AQUI"      # Sua senha real

# (v35) ATENÇÃO: URL ATUALIZADA PARA O PORTAL DE PRODUÇÃO
URL_PORTAL_PREFEITURA = "https://feiradesantanaba.webiss.com.br/"
# =====================================================================================


# =====================================================================================
# FUNÇÃO DO ROBÔ MÓDULO 3 (Prefeitura) - CORRIGIDA v41
# =====================================================================================
def executar_baixa_nfs_prefeitura(cnpj_login, senha, download_path, data_inicio_str, data_fim_str):
    """
    MÓDULO 3 - Robô para baixar XMLs de NFS-e do portal WebISS (Prefeitura).
    (v41) - Alvo: Executa o fluxo completo de navegação de menus,
                 usando o ID correto "recurso-issqn-nfse".
    """
    print(">>> [ROBÔ MÓDULO 3] Iniciando consulta na Prefeitura (WebISS - PRODUÇÃO)...")

    # --- 1. PREPARAÇÃO DO AMBIENTE ---
    dsDownloadPath_normalized = os.path.normpath(download_path)
    os.makedirs(dsDownloadPath_normalized, exist_ok=True)
    print(f"Os XMLs serão salvos em: {dsDownloadPath_normalized}")

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
    
    try:
        # (v35) Usa a URL de Produção
        navegador.get(URL_PORTAL_PREFEITURA)
        wait = WebDriverWait(navegador, 15)
        
        print("Preenchendo login e senha...")
        
        # (v29) IDs Corrigidos (Login, Senha, botao-logar)
        wait.until(EC.visibility_of_element_located((By.ID, 'Login'))).send_keys(cnpj_login)
        wait.until(EC.visibility_of_element_located((By.ID, 'Senha'))).send_keys(senha)
        
        print("Clicando em 'Entrar'...")
        wait.until(EC.element_to_be_clickable((By.ID, 'botao-logar'))).click()
        
        print("Login realizado com sucesso!")

        # --- INÍCIO DA CORREÇÃO v41 ---
        # (v40) Navegação manual pelos menus
        print("Navegando para o menu 'ISSQN'...")
        xpath_menu_issqn = "//a/span[normalize-space()='ISSQN']"
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_menu_issqn))).click()
        print("Menu 'ISSQN' clicado. Aguardando sub-menu...")
        time.sleep(1) # Pausa para o menu expandir

        print("Navegando para o sub-menu 'NFS-e'...")
        # (v41) Usando o ID exato que você encontrou!
        id_menu_nfse = "recurso-issqn-nfse"
        wait.until(EC.element_to_be_clickable((By.ID, id_menu_nfse))).click()
        print("Sub-menu 'NFS-e' clicado. Aguardando sub-menu 2...")
        time.sleep(1) # Pausa para o segundo menu expandir

        # (v40) Este é o link correto para a página de "Notas Recebidas"
        print("Clicando em 'NFS-e Tomadas/Intermediadas'...")
        xpath_notas_tomadas = "//a[contains(text(), 'NFS-e Tomadas/Intermediadas')]"
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_notas_tomadas))).click()
        print("Página 'Notas Tomadas' carregada.")
        
        # (v38) Preenche o formulário de data
        print(f"Preenchendo Data Inicial: {data_inicio_str}")
        wait.until(EC.visibility_of_element_located((By.ID, 'data-emissao-inicial'))).send_keys(data_inicio_str)
        
        print(f"Preenchendo Data Final: {data_fim_str}")
        wait.until(EC.visibility_of_element_located((By.ID, 'data-emissao-final'))).send_keys(data_fim_str)
        
        print("Clicando em 'Filtrar'...")
        wait.until(EC.element_to_be_clickable((By.ID, 'buscar-por-filtro'))).click()
        print("Filtro aplicado.")
        
        # (v39) Clica no botão "Exportar"
        print("Procurando o botão 'Exportar'...")
        time.sleep(2) # Espera a tabela de resultados (mesmo vazia) carregar
        id_exportar = "botao_abrir_modal_exportacao_notas_fiscais"
        elemento_exportar = wait.until(EC.element_to_be_clickable((By.ID, id_exportar)))
        
        navegador.execute_script("arguments[0].click();", elemento_exportar)
        print("Botão 'Exportar' clicado. O pop-up (modal) deve estar visível.")
        # --- FIM DA CORREÇÃO v41 ---

        # PAUSA PARA O RECONHECIMENTO FINAL
        print("PAUSA DE 60 SEGUNDOS PARA RECONHECIMENTO MANUAL.")
        print("Por favor, inspecione o POP-UP DE EXPORTAÇÃO e encontre o botão 'Baixar XML'!")
        time.sleep(60)

    except Exception as e:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO (MÓDULO 3) !!!!!!!!!!!!!!\n{e}")
        dir_atual = os.path.dirname(os.path.abspath(__file__))
        screenshot_filename = os.path.join(dir_atual, "mod3_prefeitura_ERRO.png")
        navegador.save_screenshot(screenshot_filename)
        navegador.quit()
        return {"sucesso": False, "mensagem": str(e), "screenshot": screenshot_filename}
    
    navegador.quit()
    print("Processo Módulo 3 (Exportar) finalizado.")
    return {"sucesso": True, "mensagem": f"Clique em 'Exportar' realizado com sucesso!"}

# =====================================================================================
# INTERFACE DO DASHBOARD (Streamlit)
# (O código do dashboard 'main' não muda, fica igual ao anterior)
# =====================================================================================
def main():
    st.set_page_config(page_title="Módulo 3 - Prefeitura", page_icon="🏛️")
    st.title("🏛️ Módulo 3: Download de NFS-e (Prefeitura)")
    st.warning("⚠️ **Atenção:** Este robô acessa o portal de **PRODUÇÃO** da prefeitura. Ele foi projetado apenas para **consultar e baixar** XMLs (operações de leitura).")
    st.markdown(f"**Alvo do Robô:** `{URL_PORTAL_PREFEITURA}`")
    
    with st.form(key="prefeitura_form"):
        st.subheader("1. Dados de Login (WebISS - Produção)")
        cnpj = st.text_input("Login (Seu CPF/CNPJ real)", value=PREFEITURA_LOGIN_CNPJ)
        senha = st.text_input("Senha (Sua Senha real)", value=PREFEITURA_SENHA, type="password")
        
        st.subheader("2. Período da Consulta")
        col1, col2 = st.columns(2)
        with col1:
            data_default_inicio = datetime.date.today() - datetime.timedelta(days=30)
            data_inicio = st.date_input("Data Inicial", value=data_default_inicio)
        with col2:
            data_default_fim = datetime.date.today()
            data_fim = st.date_input("Data Final", value=data_default_fim)
            
        st.subheader("3. Configurações do Robô")
        default_path = os.path.normpath(r'C:/Faculdade/projeto-rpa/xmls_prefeitura')
        dsDownloadPath = st.text_input("Pasta para Salvar os XMLs", value=default_path)
        
        submit_button = st.form_submit_button(label="🚀 Iniciar Robô de Download (v41)") # Texto do botão atualizado

    if submit_button:
        st.markdown("---")
        st.info("Iniciando o Robô Módulo 3...")
        st.warning("Uma janela do navegador será aberta. Não mexa no mouse ou teclado.")
        
        # (v38) Converte as datas para o formato string (dd/mm/aaaa)
        data_inicio_str = data_inicio.strftime("%d/%m/%Y")
        data_fim_str = data_fim.strftime("%d/%m/%Y")
        
        try:
            with st.spinner('O robô está fazendo login, navegando, filtrando e clicando em "Exportar"...'):
                resultado = executar_baixa_nfs_prefeitura(cnpj, senha, dsDownloadPath, data_inicio_str, data_fim_str)
            
            if resultado["sucesso"]:
                st.success("Robô finalizado com sucesso!")
                st.json(resultado)
            else:
                st.error(f"O robô falhou: {resultado['mensagem']}")
                st.image("mod3_prefeitura_ERRO.png", caption="Screenshot do erro")
        
        except Exception as e:
            st.error(f"Ocorreu um erro crítico ao rodar o robô: {e}")

if __name__ == "__main__":
    main()