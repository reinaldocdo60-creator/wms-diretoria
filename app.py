import streamlit as st
import pandas as pd
import os
from datetime import datetime

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================
st.set_page_config(
    page_title="WMS - Gestão de Estoque",
    page_icon="📦",
    layout="wide"
)

# =========================================================
# GERENCIAMENTO DE USUÁRIOS E SENHAS
# =========================================================
USUARIOS = {
    "operador": {"senha": "op123", "perfil": "OPERADOR"},
    "admin": {"senha": "admin123", "perfil": "ADMIN"}
}

# Inicialização da Sessão
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = ""
if "perfil" not in st.session_state:
    st.session_state["perfil"] = ""
if "linhas_congeladas" not in st.session_state:
    st.session_state["linhas_congeladas"] = pd.DataFrame()

# =========================================================
# CARREGAMENTO DA BASE DE DADOS
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"

@st.cache_data(ttl=5)
def carregar_dados():
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str)
        except Exception:
            df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
        return df.fillna("")
    else:
        columns = ["Cod. Interno", "Cod. Fabricante", "Descrição", "Rua", "Box", "Altura", "Garantia", "Caixa", "Data Atualização"]
        return pd.DataFrame(columns=columns)

def salvar_dados(df):
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_Dados", index=False)
    st.cache_data.clear()

if "df_base" not in st.session_state:
    st.session_state["df_base"] = carregar_dados()

df_base = st.session_state["df_base"]

# =========================================================
# TELA DE LOGIN
# =========================================================
if not st.session_state["autenticado"]:
    st.title("📦 WMS - Acesso ao Sistema")
    st.subheader("🔒 Identificação do Usuário")

    col_login, _ = st.columns([1, 1])
    with col_login:
        usuario_input = st.selectbox("Selecione o Perfil / Usuário", list(USUARIOS.keys()))
        senha_input = st.text_input("Senha de Acesso", type="password")

        if st.button("🔑 Entrar no WMS", use_container_width=True):
            user_info = USUARIOS.get(usuario_input.lower())
            if user_info and user_info["senha"] == senha_input:
                st.session_state["autenticado"] = True
                st.session_state["usuario_logado"] = usuario_input.capitalize()
                st.session_state["perfil"] = user_info["perfil"]
                st.rerun()
            else:
                st.error("❌ Senha incorreta!")
    st.stop()

# =========================================================
# BARRA LATERAL (NAV / PERFIL)
# =========================================================
st.sidebar.title("📦 WMS Nuvem")
st.sidebar.write(f"👤 **Usuário:** {st.session_state['usuario_logado']}")
st.sidebar.write(f"🛡️ **Perfil:** `{st.session_state['perfil']}`")

st.sidebar.markdown("---")

opcoes_menu = [
    "🔍 Pesquisa e Validação (Geral)",
    "🚚 Mover Produto",
    "➕ Cadastrar / Ocupar",
    "🧹 Limpar Endereço",
    "📥 Importar / Atualizar Base em Massa",
    "💾 Backup e Histórico"
]

opcao_menu = st.sidebar.radio("Navegação Principal", opcoes_menu)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Sair do Sistema"):
    st.session_state["autenticado"] = False
    st.session_state["usuario_logado"] = ""
    st.session_state["perfil"] = ""
    st.rerun()

def validar_admin():
    if st.session_state["perfil"] != "ADMIN":
        st.error("⚠️ Acesso restrito! Esta funcionalidade exige perfil de ADMINISTRADOR.")
        return False
    return True

# =========================================================
# TELA 1: PESQUISA E VALIDAÇÃO (GERAL)
# =========================================================
if opcao_menu == "🔍 Pesquisa e Validação (Geral)":
    st.header("🔍 Pesquisa, Validação e Congelamento de Endereços")
    st.subheader("📌 Painel de Operação")

    q_busca = st.text_input("1️⃣ PESQUISA: Digite Cód. Interno, Fabricante, Descrição, Rua, Box ou Altura:").strip().upper()
    q_valid = st.text_input("2️⃣ VALIDAÇÃO: Bipe ou digite o Cód. Fabricante para conferência:").strip().upper()

    df_res = df_base.copy()
    if q_busca:
        mask = (
            df_res["Cod. Interno"].astype(str).str.upper().str.contains(q_busca) |
            df_res["Cod. Fabricante"].astype(str).str.upper().str.contains(q_busca) |
            df_res["Descrição"].astype(str).str.upper().str.contains(q_busca) |
            df_res["Rua"].astype(str).str.upper().str.contains(q_busca) |
            df_res["Box"].astype(str).str.upper().str.contains(q_busca) |
            df_res["Altura"].astype(str).str.upper().str.contains(q_busca)
        )
        df_res = df_res[mask]

    col_btn1, col_btn2, col_btn3 = st.columns(3)
    with col_btn1:
        if st.button("❄️ CONGELAR LINHAS DA PESQUISA", use_container_width=True):
            if not df_res.empty:
                st.session_state["linhas_congeladas"] = pd.concat([st.session_state["linhas_congeladas"], df_res]).drop_duplicates()
                st.success("Linhas congeladas com sucesso!")
    with col_btn2:
        if st.button("🔥 DESCONGELAR / LIMPAR ACÚMULO", use_container_width=True):
            st.session_state["linhas_congeladas"] = pd.DataFrame()
            st.info("Acúmulo congelado limpo!")
    with col_btn3:
        st.write("🖨️ Para imprimir o acúmulo use `Ctrl + P`")

    tab1, tab2 = st.tabs([f"🔎 Resultado da Busca ({len(df_res)})", f"❄️ Linhas Acumuladas / Congeladas ({len(st.session_state['linhas_congeladas'])})"])

    with tab1:
        st.dataframe(df_res, use_container_width=True)

    with tab2:
        st.dataframe(st.session_state["linhas_congeladas"], use_container_width=True)

    if q_valid:
        if not df_res.empty and (df_res["Cod. Fabricante"].astype(str).str.upper() == q_valid).any():
            st.success(f"✅ VALIDAÇÃO OK: Código {q_valid} pertence à busca realizada!")
        else:
            st.error(f"❌ ATENÇÃO: Código {q_valid} NÃO ENCONTRADO na lista atual!")

# =========================================================
# TELA 2: MOVER PRODUTO
# =========================================================
elif opcao_menu == "🚚 Mover Produto":
    st.header("🚚 Movimentação Interna de Produto")
    
    with st.form("form_mover"):
        cod_mover = st.text_input("Código do Produto (Interno ou Fabricante) *").strip().upper()
        nova_rua = st.text_input("Nova Rua *").strip().upper()
        novo_box = st.text_input("Novo Box *").strip().upper()
        nova_altura = st.text_input("Nova Altura *").strip().upper()
        
        btn_mover = st.form_submit_button("Confirmar Movimentação")
        
        if btn_mover:
            if not cod_mover or not nova_rua or not novo_box or not nova_altura:
                st.warning("Preencha todos os campos obrigatórios (*).")
            else:
                idx = df_base[(df_base["Cod. Interno"].str.upper() == cod_mover) | (df_base["Cod. Fabricante"].str.upper() == cod_mover)].index
                if not idx.empty:
                    df_base.loc[idx, "Rua"] = nova_rua
                    df_base.loc[idx, "Box"] = novo_box
                    df_base.loc[idx, "Altura"] = nova_altura
                    df_base.loc[idx, "Data Atualização"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    salvar_dados(df_base)
                    st.success("✅ Produto movimentado com sucesso!")
                else:
                    st.error("Produto não localizado no estoque.")

# =========================================================
# TELA 3: CADASTRAR / OCUPAR
# =========================================================
elif opcao_menu == "➕ Cadastrar / Ocupar":
    st.header("➕ Cadastrar / Ocupar Endereço")
    if validar_admin():
        with st.form("form_cadastrar"):
            c1, c2 = st.columns(2)
            cod_int = c1.text_input("Cód. Interno *").strip().upper()
            cod_fab = c2.text_input("Cód. Fabricante *").strip().upper()
            desc = st.text_input("Descrição Completa *").strip().upper()
            
            e1, e2, e3 = st.columns(3)
            rua = e1.text_input("Rua *").strip().upper()
            box = e2.text_input("Box *").strip().upper()
            altura = e3.text_input("Altura *").strip().upper()
            
            btn_cad = st.form_submit_button("Cadastrar / Ocupar")
            if btn_cad:
                if not (cod_int and cod_fab and desc and rua and box and altura):
                    st.warning("Preencha todos os campos obrigatórios (*).")
                else:
                    novo_registro = {
                        "Cod. Interno": cod_int,
                        "Cod. Fabricante": cod_fab,
                        "Descrição": desc,
                        "Rua": rua,
                        "Box": box,
                        "Altura": altura,
                        "Garantia": "",
                        "Caixa": "",
                        "Data Atualização": datetime.now().strftime("%Y-%m-%d %H:%M")
                    }
                    df_base = pd.concat([df_base, pd.DataFrame([novo_registro])], ignore_index=True)
                    salvar_dados(df_base)
                    st.session_state["df_base"] = df_base
                    st.success("✅ Novo produto cadastrado/endereçado com sucesso!")

# =========================================================
# TELA 4: LIMPAR ENDEREÇO
# =========================================================
elif opcao_menu == "🧹 Limpar Endereço":
    st.header("🧹 Desocupar / Limpar Endereço")
    if validar_admin():
        with st.form("form_limpar"):
            cod_limp = st.text_input("Código Interno ou Fabricante a Desocupar *").strip().upper()
            btn_limp = st.form_submit_button("Desocupar Endereço")
            
            if btn_limp:
                idx = df_base[(df_base["Cod. Interno"].str.upper() == cod_limp) | (df_base["Cod. Fabricante"].str.upper() == cod_limp)].index
                if not idx.empty:
                    df_base.loc[idx, ["Rua", "Box", "Altura"]] = ""
                    df_base.loc[idx, "Data Atualização"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    salvar_dados(df_base)
                    st.session_state["df_base"] = df_base
                    st.success("✅ Endereço desocupado com sucesso!")
                else:
                    st.error("Produto não localizado no estoque.")

# =========================================================
# TELA 5: IMPORTAÇÃO / ATUALIZAÇÃO DA BASE EM MASSA (ADMIN)
# =========================================================
elif opcao_menu == "📥 Importar / Atualizar Base em Massa":
    st.header("📥 Importação e Atualização da Base em Massa")
    if validar_admin():
        st.write("Suba o arquivo Excel (`.xlsx`) com o mapeamento completo do estoque para atualizar a base de dados do WMS em tempo real durante a reunião.")
        
        arquivo_enviado = st.file_uploader("Selecione a planilha Excel mapeada", type=["xlsx"])
        
        if arquivo_enviado is not None:
            try:
                try:
                    df_novo = pd.read_excel(arquivo_enviado, sheet_name="Base_Dados", dtype=str)
                except Exception:
                    df_novo = pd.read_excel(arquivo_enviado, dtype=str)
                
                df_novo = df_novo.fillna("")
                
                st.success("Planilha lida com sucesso! Pré-visualização das 10 primeiras linhas:")
                st.dataframe(df_novo.head(10), use_container_width=True)
                
                if st.button("🚀 Confirmar e Atualizar Base do WMS", use_container_width=True):
                    salvar_dados(df_novo)
                    st.session_state["df_base"] = df_novo
                    st.balloons()
                    st.success("✅ Base de dados do WMS atualizada com sucesso! O novo mapeamento já está ativo para consulta e bipagem.")
            except Exception as e:
                st.error(f"Erro ao processar o arquivo Excel: {e}")

# =========================================================
# TELA 6: BACKUP E HISTÓRICO
# =========================================================
elif opcao_menu == "💾 Backup e Histórico":
    st.header("💾 Backup e Exportação da Base")
    st.write("Baixe uma cópia da base de estoque atual em Excel:")
    
    if os.path.exists(ARQUIVO_EXCEL):
        with open(ARQUIVO_EXCEL, "rb") as f:
            st.download_button(
                label="📥 Baixar Base_Estoque.xlsx Atualizada",
                data=f,
                file_name=f"Backup_WMS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
