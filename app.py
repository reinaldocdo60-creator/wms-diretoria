import streamlit as str_lit
import pandas as pd
import os
import json
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================
# CONFIGURAÇÃO DO GOOGLE GEMINI E AMBIENTE
# =====================================================
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Carrega as variáveis de ambiente locais (se houver arquivo .env)
load_dotenv()

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
        /* Oculta elementos desnecessários da interface do Streamlit na impressão */
        header, footer, nav, .stSidebar, [data-testid="stSidebar"], .stButton, .stDownloadButton {
            display: none !important;
        }
        
        /* Ajusta o layout principal para ocupar a página inteira */
        body, .main, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
            background-color: white !important;
            color: black !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        
        /* Garante que tabelas e textos fiquem legíveis em preto e branco */
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
    
    # Exemplo de conteúdo da tela de estoque
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

    # Inicializa o cliente do Gemini de forma robusta
    try:
        # Tenta inicializar usando a chave informada diretamente ou via ambiente
        client = genai.Client(api_key="AQ.Ab8RN6I1bXpNqsD3J7zN046MfeH4ld3DmrwraC1srEuEAnczaA")
    except Exception as e:
        try:
            client = genai.Client()
        except Exception as e2:
            client = None

    if not client:
        str_lit.error("Erro ao inicializar a IA. Verifique se a chave de API está correta.")

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

        # Processa a resposta com o Gemini
        if client:
            with str_lit.chat_message("assistant"):
                with str_lit.spinner("Consultando as regras do WMS..."):
                    try:
                        instrucao_sistema = """
                        Você é o assistente virtual oficial de um Sistema de Gestão de Armazém (WMS).
                        Responda dúvidas sobre as rotinas, processos, controle de estoque e regras de negócio do sistema de forma clara, prestativa e em português brasileiro.
                        """

                        response = client.models.generate_content(
                            model="gemini-2.5-flash", 
                            contents=duvida_usuario,
                            config=types.GenerateContentConfig(
                                system_instruction=instrucao_sistema,
                                temperature=0.3,
                            ),
                        )
                        
                        resposta_ia = response.text
                        str_lit.markdown(resposta_ia)
                        
                        # Adiciona a resposta da IA ao histórico
                        str_lit.session_state["historico_chat"].append({"role": "assistant", "content": resposta_ia})

                    except Exception as erro:
                        str_lit.error(f"Desculpe, ocorreu um erro ao consultar a IA: {erro}")
        else:
            str_lit.warning("A IA está desativada temporariamente porque a chave da API não foi encontrada.")
