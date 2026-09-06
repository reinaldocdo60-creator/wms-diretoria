import streamlit as st
import pandas as pd
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="WMS - Gestão de Estoque",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CONFIGURAÇÃO DE USUÁRIOS INICIAIS ---
if "usuarios_db" not in st.session_state:
    st.session_state["usuarios_db"] = {
        "admin": {"senha": "123", "perfil": "ADMIN"},
        "operador": {"senha": "123", "perfil": "OPERADOR"}
    }

# --- TELA DE LOGIN ---
def tela_login():
    st.sidebar.title("🔐 Acesso ao WMS")
    usuario_input = st.sidebar.text_input("Usuário").strip().lower()
    senha_input = st.sidebar.text_input("Senha", type="password")
    
    if st.sidebar.button("Entrar", use_container_width=True):
        db = st.session_state["usuarios_db"]
        if usuario_input in db and db[usuario_input]["senha"] == senha_input:
            st.session_state["autenticado"] = True
            st.session_state["usuario_atual"] = usuario_input
            st.session_state["perfil_atual"] = db[usuario_input]["perfil"]
            st.rerun()
        else:
            st.sidebar.error("Usuário ou senha incorretos.")

# --- VERIFICAÇÃO DE AUTENTICAÇÃO ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("📦 Sistema WMS - Login Necessário")
    st.info("Por favor, faça login na barra lateral para acessar o sistema.")
    tela_login()
    st.stop()

# --- MENU LATERAL APÓS LOGIN ---
st.sidebar.success(f"Logado: {st.session_state['usuario_atual']} ({st.session_state['perfil_atual']})")

menu = ["Consulta / Operação", "Atualizar Base (Admin)"]

# Se for Admin, adicionamos a opção de gerenciar usuários
if st.session_state["perfil_atual"] == "ADMIN":
    menu.append("Gerenciar Usuários")

menu.append("Sair")
escolha = st.sidebar.selectbox("Menu de Navegação", menu)

# Ação de Sair
if escolha == "Sair":
    st.session_state["autenticado"] = False
    st.rerun()

# --- CARREGAMENTO DA BASE DE DADOS (ARQUIVO ÚNICO) ---
ARQUIVO_BASE = "Base_Estoque.xlsx"

def carregar_dados():
    if os.path.exists(ARQUIVO_BASE):
        try:
            return pd.read_excel(ARQUIVO_BASE)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo Excel: {e}")
            return None
    return None

# --- ROTEAMENTO DAS TELAS ---

if escolha == "Consulta / Operação":
    st.title("🔍 Consulta de Estoque e Operação")
    
    df = carregar_dados()
    if df is not None and not df.empty:
        # Campo de pesquisa rápida
        pesquisa = st.text_input("🔍 Pesquisar por Código, Descrição, Endereço ou Fabricante:")
        
        if pesquisa:
            # Filtro inteligente em todas as colunas convertidas para texto
            mask = df.astype(str).apply(lambda x: x.str.contains(pesquisa, case=False, na=False)).any(axis=1)
            df_resultado = df[mask]
            st.info(f"Encontrados {len(df_resultado)} registros para a pesquisa.")
            st.dataframe(df_resultado, use_container_width=True)
        else:
            st.write(f"Exibindo visão geral do estoque ({len(df)} itens totais):")
            st.dataframe(df, use_container_width=True)
    else:
        st.warning(f"O arquivo '{ARQUIVO_BASE}' não foi encontrado ou está vazio. Vá até 'Atualizar Base (Admin)' para enviar a planilha.")

elif escolha == "Atualizar Base (Admin)":
    if st.session_state["perfil_atual"] != "ADMIN":
        st.error("Acesso negado. Apenas administradores podem atualizar a base de dados.")
    else:
        st.title("📂 Atualização em Massa da Base")
        st.write("Faça o upload do novo arquivo Excel (`Base_Estoque.xlsx`) para atualizar o estoque com segurança.")
        
        uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=["xlsx", "xls"])
        if uploaded_file is not None:
            try:
                with open(ARQUIVO_BASE, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.success("Base de estoque atualizada com sucesso no servidor! As alterações já estão valendo para a operação.")
            except Exception as e:
                st.error(f"Erro ao salvar o arquivo: {e}")

elif escolha == "Gerenciar Usuários":
    if st.session_state["perfil_atual"] != "ADMIN":
        st.error("Acesso negado.")
    else:
        st.title("👥 Gerenciamento de Colaboradores e Acessos")
        st.write("Adicione novos operadores ou remova acessos diretamente por aqui.")
        
        db = st.session_state["usuarios_db"]
        
        # Exibir usuários atuais em uma tabela limpa
        dados_tabela = [{"Usuário": k, "Perfil": v["perfil"]} for k, v in db.items()]
        st.dataframe(pd.DataFrame(dados_tabela), use_container_width=True)
        
        st.divider()
        
        st.subheader("Cadastrar Novo Colaborador")
        col1, col2, col3 = st.columns(3)
        with col1:
            novo_usuario = st.text_input("Nome de Usuário (ex: joao.silva)").strip().lower()
        with col2:
            nova_senha = st.text_input("Senha Inicial", type="password")
        with col3:
            novo_perfil = st.selectbox("Perfil de Acesso", ["OPERADOR", "ADMIN"])
        
        if st.button("Cadastrar Usuário", use_container_width=True):
            if not novo_usuario or not nova_senha:
                st.warning("Preencha o usuário e a senha.")
            elif novo_usuario in db:
                st.error("Este usuário já existe!")
            else:
                st.session_state["usuarios_db"][novo_usuario] = {
                    "senha": nova_senha,
                    "perfil": novo_perfil
                }
                st.success(f"Usuário '{novo_usuario}' cadastrado com sucesso!")
                st.rerun()
                
        st.divider()
        st.subheader("Remover Usuário")
        
        usuarios_removiveis = [u for u in db.keys() if u != "admin"]
        
        if usuarios_removiveis:
            usuario_para_remover = st.selectbox("Selecione o usuário para excluir", usuarios_removiveis)
            if st.button("Excluir Usuário Selecionado", type="primary"):
                if usuario_para_remover in st.session_state["usuarios_db"]:
                    del st.session_state["usuarios_db"][usuario_para_remover]
                    st.success(f"Usuário '{usuario_para_remover}' removido com sucesso!")
                    st.rerun()
        else:
            st.info("Não há outros usuários cadastrados para remoção além do admin principal.")
