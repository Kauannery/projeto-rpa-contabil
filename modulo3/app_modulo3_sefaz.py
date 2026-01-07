#!/usr/bin/env python
# NOME DESTE ARQUIVO: app_modulo3_sefaz.py
#
# Este arquivo é o Módulo 3 do TCC, focado em demonstrar
# o portal de alta segurança da SEFAZ (e-CAC).

# --- Importações ---
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import sys

# --- Importações para o Dashboard (Streamlit) ---
import streamlit as st

# =====================================================================================
# INSTRUÇÕES DESTE MÓDULO
# =====================================================================================
"""
## 1. Objetivo deste Módulo

Este robô demonstra o acesso ao portal de alta segurança (e-CAC) da Receita Federal,
onde é realizada a consulta de NF-e (produtos).

Este portal utiliza autenticação via Certificado Digital, um mecanismo de segurança
nível bancário que o RPA de interface (Selenium) não manipula diretamente.

    
"""

# =====================================================================================
# FUNÇÃO DO ROBÔ MÓDULO 3 (SEFAZ NF-e)
# =====================================================================================
def executar_demonstracao_sefaz_nfe():
    """
    MÓDULO 3 - Robô de demonstração do portal e-CAC (Certificado Digital).
    Navega até o portal e-CAC para demonstrar o processo de login.
    """
    print(">>> [ROBÔ MÓDULO 3] Iniciando demonstração no e-CAC...")

    # --- Configuração do Chrome ---
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--window-size=1024,768')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--ignore-certificate-errors')

    navegador = webdriver.Chrome(options=chrome_options)
    
    try:
        # 1. Navega até o "cofre" da Receita Federal
        url_ecac = "https://cav.receita.fazenda.gov.br/autenticacao/login"
        print(f"Acessando o portal e-CAC: {url_ecac}")
        navegador.get(url_ecac)
        
        wait = WebDriverWait(navegador, 15)
        
        # 2. Espera o botão "Certificado Digital" aparecer
        print("Aguardando a tela de login (portal do Certificado)...")
        xpath_botao_certificado = "//*[@id='cert-digital-icon']"
        wait.until(EC.visibility_of_element_located((By.XPATH, xpath_botao_certificado)))
        
        print("Portal de alta segurança e-CAC carregado!")
        
        # 3. Pausa para a demonstração (para a banca do TCC ver)
        print("Robô pausado por 15 segundos para demonstração da tela...")
        time.sleep(15)

    except Exception as e:
        print(f"\n!!!!!!!!!!!!!!!! OCORREU UM ERRO (MÓDULO 3) !!!!!!!!!!!!!!\n{e}")
        dir_atual = os.path.dirname(os.path.abspath(__file__))
        screenshot_filename = os.path.join(dir_atual, "mod3_sefaz_ERRO.png")
        navegador.save_screenshot(screenshot_filename)
        navegador.quit()
        return {"sucesso": False, "mensagem": str(e), "screenshot": screenshot_filename}
    
    navegador.quit()
    print("Demonstração do Módulo 3 finalizada.")
    
    # --- (v29) MENSAGEM DO DASHBOARD ATUALIZADA (NEUTRA) ---
    return {"sucesso": True, "mensagem": "Demonstração concluída. O robô acessou com sucesso o portal e-CAC, que utiliza autenticação de alta segurança (Certificado Digital) para o acesso aos serviços de NF-e."}

# =====================================================================================
# INTERFACE DO DASHBOARD (Streamlit) - MÓDULO 3 (v29 - TEXTO LIMPO)
# =====================================================================================
def main():
    st.set_page_config(page_title="Módulo 3 - SEFAZ NF-e", page_icon="🛡️")
    
    # --- (v29) Título e Textos Atualizados ---
    st.title("🛡️ Módulo 3: Acesso ao Portal SEFAZ (NF-e)")
    st.info("Este robô demonstra o acesso ao portal de alta segurança (e-CAC) da Receita Federal, onde é realizada a consulta de NF-e (produtos).")
    st.markdown("---")
    
    st.subheader("Objetivo da Demonstração:")
    st.markdown("""
    1.  O robô irá acessar o **e-CAC** (Centro Virtual de Atendimento) da Receita Federal.
    2.  Este portal centraliza os serviços de alta segurança, como o download em lote de NF-e.
    3.  A automação irá navegar e exibir a tela de login principal, que utiliza **Certificado Digital**.
    
    *(Esta demonstração conclui a análise de viabilidade do RPA de interface (Selenium) para este tipo de portal, indicando que a automação completa exigiria um robô de backend com acesso direto à API/WebService do governo, como discutido na documentação do TCC.)*
    """)
    
    submit_button = st.button(label="🚀 Iniciar Demonstração de Acesso (e-CAC)")

    if submit_button:
        st.markdown("---")
        st.info("Iniciando o Robô Módulo 3...")
        st.warning("Uma janela do navegador será aberta. Observe a tela de login do e-CAC.")
        
        try:
            with st.spinner('O robô está navegando até o portal e-CAC...'):
                resultado = executar_demonstracao_sefaz_nfe()
            
            if resultado["sucesso"]:
                # (v29) Mensagem de Sucesso Neutra
                st.success("Demonstração Concluída com Sucesso!")
                st.markdown(f"**Resultado:** {resultado['mensagem']}")
            else:
                st.error(f"O robô falhou: {resultado['mensagem']}")
                st.image("mod3_sefaz_ERRO.png", caption="Screenshot do erro")
        
        except Exception as e:
            st.error(f"Ocorreu um erro crítico ao rodar o robô: {e}")

if __name__ == "__main__":
    main()