import streamlit as str_lit
import pandas as pd
import os
import json
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================
# CONFIGURAÇÃO DO GOOGLE GEMINI (Biblioteca Clássica)
# =====================================================
import google.generativeai as genai
from dotenv import load_dotenv

# Carrega as variáveis de ambiente locais (se houver arquivo .env)
load_dotenv()

# Configura a API key diretamente com a chave informada
CHAVE_API_WMS = "AQ.Ab8RN6I1bXpNqsD3J7zN046MfeH4ld3DmrwraC1srEuEAnczaA"
genai.configure(api_key=CHAVE_API_WMS)

# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
str_lit.set_page_config(
    page_title="WMS - Gestão de Estoque",
    page_icon="📦",
    layout="wide"
)

# =====================================================
# ESTILO CSS GLOBAL + REGRAS DE IMPRESSÃO (MEDIA PRINT)
# =====================================================
str_lit.markdown("""
<style>
    /* Estilos gerais da aplicação */
    .main {
        background-color: #f8f9fa;
    }
    
    /* Regras estritas para impressão e exportação para PDF limpa */
    @media print {
        header, footer, nav, .stSidebar, [data-testid="stSidebar"], .stButton, .stDownloadButton {
            display: none !important;
        }
        body, .main, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
            background-color: white !important;
            color: black !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        table {
            border-collapse: collapse !important;
            width: 100% !important;
        }
        th, td {
            border: 1px solid #ddd !important;
            padding: 8px !important;
            color: black !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# MENU DE NAVEGAÇÃO LATERAL (Com a opção do Assistente IA)
# =====================================================
opcao_menu = str_lit.sidebar.selectbox("Navegação", ["Estoque", "Relatórios", "🤖 Assistente IA"])

# =====================================================
# TELA 1: ESTOQUE
# =====================================================
if opcao_menu == "Estoque":
    str_lit.title("📦 WMS - Gestão de Estoque")
    str_lit.markdown("Gerencie os produtos, entradas e saídas do armazém.")
    
    try:
        df_estoque = pd.read_excel("Base_Estoque.xlsx")
        str_lit.dataframe(df_estoque, use_container_width=True)
    except Exception:
        str_lit.info("Base de estoque carregada ou aguardando arquivo de dados.")

# =====================================================
# TELA 2: RELATÓRIOS
# =====================================================
elif opcao_menu == "Relatórios":
    str_lit.title("📊 Relatórios e Indicadores")
    str_lit.markdown("Visualize as métricas de movimentação e inventário.")
    str_lit.write("Painel de relatórios em andamento...")

# =====================================================
# TELA 3: ASSISTENTE VIRTUAL DE IA (GEMINI)
# =====================================================
elif opcao_menu == "🤖 Assistente IA":
    str_lit.title("🤖 Assistente Virtual WMS")
    str_lit.markdown("Tire suas dúvidas sobre as rotinas, processos e regras de negócio do nosso sistema de gerenciamento de armazém.")

    # Histórico de mensagens do chat na sessão do Streamlit
    if "historico_chat" not in str_lit.session_state:
        str_lit.session_state["historico_chat"] = []

    # Exibe as mensagens anteriores do chat
    for mensagem in str_lit.session_state["historico_chat"]:
        with str_lit.chat_message(mensagem["role"]):
            str_lit.markdown(mensagem["content"])

    # Caixa de texto flutuante para entrada do usuário
    if duvida_usuario := str_lit.chat_input("Digite sua dúvida sobre o WMS ou operações de armazém..."):
        # Adiciona a mensagem do usuário ao histórico
        str_lit.session_state["historico_chat"].append({"role": "user", "content": duvida_usuario})
        with str_lit.chat_message("user"):
            str_lit.markdown(duvida_usuario)

        # Processa a resposta com o Gemini clássico
        with str_lit.chat_message("assistant"):
            with str_lit.spinner("Consultando as regras do WMS..."):
                try:
                    instrucao_sistema = """
                    Você é o assistente virtual oficial de um Sistema de Gestão de Armazém (WMS).
                    Responda dúvidas sobre as rotinas, processos, controle de estoque e regras de negócio do sistema de forma clara, prestativa e em português brasileiro.
                    """

                    model = genai.GenerativeModel(
                        model_name="gemini-1.5-flash",
                        system_instruction=instrucao_sistema
                    )
                    
                    response = model.generate_content(duvida_usuario)
                    
                    resposta_ia = response.text
                    str_lit.markdown(resposta_ia)
                    
                    # Adiciona a resposta da IA ao histórico
                    str_lit.session_state["historico_chat"].append({"role": "assistant", "content": resposta_ia})

                except Exception as erro:
                    str_lit.error(f"Desculpe, ocorreu um erro ao consultar a IA: {erro}")
