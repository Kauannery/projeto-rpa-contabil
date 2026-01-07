# Protótipo RPA Contábil - TCC (Módulo: Emissão de NFS-e)

Este projeto é um protótipo de RPA (Robotic Process Automation) focado em automação de tarefas contábeis, desenvolvido como parte de um Trabalho de Conclusão de Curso (TCC).

Este módulo específico (`emissornfe`) automatiza o processo de emissão de Nota Fiscal de Serviço eletrônica (NFS-e) para MEI (Microempreendedor Individual) no portal nacional de produção restrita (ambiente de teste).

## 1. Estrutura dos Arquivos

O projeto nesta pasta está dividido em dois arquivos Python principais:

1.  `emissor_rpa.py`: Este é o "motor" do robô. É um módulo Python que contém toda a lógica de automação do Selenium para navegar no portal do governo, preencher os dados e emitir a nota fiscal. Ele é feito para ser "chamado" por uma interface.
2.  `dashboard.py`: Este é o "painel de controle" (a interface do usuário). É um aplicativo web criado com a biblioteca Streamlit. Ele fornece um formulário amigável para o usuário inserir os dados e um botão para "chamar" o robô `emissor_rpa.py`.

## 2. Pré-requisitos (O que você precisa ter instalado)

Antes de rodar o projeto, garanta que você tenha os seguintes itens instalados no seu computador:

1.  **Python 3.x**: [python.org](https://www.python.org/)
    * *Durante a instalação, marque a caixa "Add Python to PATH"*.
2.  **Google Chrome**: O navegador que o robô irá controlar.

## 3. Instalação (Comandos para o Terminal)

Abra seu terminal (CMD ou PowerShell) e instale as bibliotecas Python necessárias para o projeto:

```bash
# Instalar a biblioteca de automação web (o robô)
pip install selenium

# Instalar a biblioteca do dashboard (a interface web)
pip install streamlit

# Navegue até a pasta exata do projeto:

cd C:\Faculdade\projeto-rpa\emissornfe

# Execute o seguinte comando do Streamlit:
python -m streamlit run dashboard.py
