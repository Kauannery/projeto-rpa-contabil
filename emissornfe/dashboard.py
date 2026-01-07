#!/usr/bin/env python
# NOME DESTE ARQUIVO: dashboard.py

import streamlit as st
import os

# Importa a função do nosso outro arquivo!
# Isso só funciona porque 'emissor_rpa.py' está na mesma pasta.
from emissor_rpa import executar_emissao

# --- Configuração da Página ---
st.set_page_config(
    page_title="RPA Contábil - TCC",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="auto"
)

# --- Título e Descrição ---
st.title("🤖 Protótipo RPA Contábil - TCC")
st.markdown("Bem-vindo ao dashboard de automação. Este protótipo utiliza RPA (Robotic Process Automation) para executar tarefas contábeis.")
st.header("Módulo 1: Emissão de NFS-e (MEI)")
st.info("Preencha os dados abaixo e clique em 'Emitir Nota' para iniciar o robô. A emissão será feita no **Ambiente de Teste** do governo (Produção Restrita).")

# --- Formulário de Entrada de Dados ---
with st.form(key="emissao_form"):
    st.subheader("1. Dados do Prestador (Você - MEI)")
    # Usamos colunas para organizar o formulário
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

    st.subheader("4. Configurações do Robô")
    # Define o caminho padrão de download
    default_path = os.path.normpath(r'C:/Faculdade/projeto-rpa/notas_teste')
    dsDownloadPath = st.text_input("Pasta para Salvar as Notas", value=default_path)
    
    inTerminal = st.checkbox("Rodar robô em segundo plano (headless)", value=False)
    
    # Botão de envio do formulário
    submit_button = st.form_submit_button(label="🚀 Emitir Nota Fiscal (Teste)")


# --- Lógica de Execução ---
if submit_button:
    # Quando o botão for clicado, o código aqui é executado
    st.markdown("---")
    st.info("O robô foi iniciado! Por favor, aguarde...")
    st.warning("Uma janela do navegador será aberta. Não mexa no mouse ou teclado até o processo terminar.")
    
    # 1. Monta o dicionário de configuração com os dados do formulário
    config_rpa = {
        "dsEmissorCNPJ": dsEmissorCNPJ,
        "dsEmissorPass": dsEmissorPass,
        "vlNota": vlNota.replace(',', '.'), # Garante o formato correto do valor
        "cdTomador": cdTomador,
        "dsTomadorCEP": dsTomadorCEP,
        "dsBuscaMunicipio": "São Paulo", # Valor fixo por enquanto
        "dsMunicipio": "São Paulo/SP",   # Valor fixo por enquanto
        "dsTributario": dsTributario,
        "dsServico": dsServico,
        "dsDownloadPath": dsDownloadPath,
        "inTerminal": inTerminal
    }
    
    # 2. Chama a função do nosso robô!
    try:
        with st.spinner('O robô está trabalhando... (Isso pode levar 1-2 minutos)'):
            resultado = executar_emissao(config_rpa)
        
        # 3. Mostra o resultado na tela
        if resultado["sucesso"]:
            st.success("Robô finalizado com sucesso!")
            st.balloons() # Comemoração!
            st.json(resultado)
        else:
            st.error(f"O robô falhou: {resultado['mensagem']}")
            if "screenshot" in resultado:
                st.image("emissor_ERRO.png", caption=f"Screenshot do erro: {resultado['screenshot']}")
                
    except Exception as e:
        st.error(f"Ocorreu um erro crítico ao tentar rodar o robô: {e}")