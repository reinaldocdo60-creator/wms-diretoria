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
# CARREGAMENTO SEGURO DA BASE DE DADOS (BLINDADO PARA MOBILE)
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"

@st.cache_data(ttl=2)
def carregar_dados():
    colunas_padrao = ["Cod. Interno", "Cod. Fabricante", "Descrição", "Rua", "Box", "Altura", "Garantia", "Caixa", "Data Atualização"]
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            try:
                df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str)
            except Exception:
                df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
            
            df = df.fillna("")
            for col in colunas_padrao:
                if col not in df.columns:
                    df[col] = ""
            return df
        except Exception:
            return pd.DataFrame(columns=colunas_padrao)
    else:
        return pd.DataFrame(columns=colunas_padrao)

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
# BARRA LATERAL (MENU E PERFIL)
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
if st.sidebar.button("🚪 Sair do Sistema", use_container_width=True):
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
# TELA 1: PESQUISA E VALIDAÇÃO (OTIMIZADA PARA CELULAR)
# =========================================================
if opcao_menu == "🔍 Pesquisa e Validação (Geral)":
    st.header("🔍 Pesquisa e Validação")

    q_busca = st.text_input("1️⃣ PESQUISA (Cód., Descrição, Endereço):", key="q_busca").strip().upper()
    q_valid = st.text_input("2️⃣ VALIDAÇÃO (Bipe o Cód. Fabricante):", key="q_valid").strip().upper()

    df_res = st.session_state.get("df_base", carregar_dados()).copy()

    if q_busca and not df_res.empty:
        try:
            mask = (
                df_res["Cod. Interno"].astype(str).str.upper().str.contains(q_busca, regex=False) |
                df_res["Cod. Fabricante"].astype(str).str.upper().str.contains(q_busca, regex=False) |
                df_res["Descrição"].astype(str).str.upper().str.contains(q_busca, regex=False) |
                df_res["Rua"].astype(str).str.upper().str.contains(q_busca, regex=False) |
                df_res["Box"].astype(str).str.upper().str.contains(q_busca, regex=False) |
                df_res["Altura"].astype(str).str.upper().str.contains(q_busca, regex=False)
            )
            df_res = df_res[mask]
        except Exception as e:
            st.error(f"Erro ao filtrar busca: {e}")

    if st.button("❄️ CONGELAR LINHAS DA PESQUISA", use_container_width=True):
        if not df_res.empty:
            congeladas_atuais = st.session_state.get("linhas_congeladas", pd.DataFrame())
            st.session_state["linhas_congeladas"] = pd.concat([congeladas_atuais, df_res]).drop_duplicates()
            st.success("Linhas congeladas com sucesso!")

    if st.button("🔥 LIMPAR LINHAS CONGELADAS", use_container_width=True):
        st.session_state["linhas_congeladas"] = pd.DataFrame()
        st.info("Acúmulo de linhas limpo!")

    tab1, tab2 = st.tabs([f"🔎 Resultado ({len(df_res)})", f"❄️ Congeladas ({len(st.session_state.get('linhas_congeladas', pd.DataFrame()))})"])

    with tab1:
        st.dataframe(df_res, use_container_width=True)

    with tab2:
        st.dataframe(st.session_state.get("linhas_congeladas", pd.DataFrame()), use_container_width=True)

    if q_valid:
        if not df_res.empty and (df_res["Cod. Fabricante"].astype(str).str.upper() == q_valid).any():
            st.success(f"✅ VALIDAÇÃO OK: Código {q_valid} pertence à busca!")
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
        
        btn_mover = st.form_submit_button("Confirmar Movimentação", use_container_width=True)
        
        if btn_mover:
            if not cod_mover or not nova_rua or not novo_box or not nova_altura:
                st.warning("Preencha todos os campos obrigatórios (*).")
            else:
                df_atual = st.session_state["df_base"]
                idx = df_atual[(df_atual["Cod. Interno"].str.upper() == cod_mover) | (df_atual["Cod. Fabricante"].str.upper() == cod_mover)].index
                if not idx.empty:
                    df_atual.loc[idx, "Rua"] = nova_rua
                    df_atual.loc[idx, "Box"] = novo_box
                    df_atual.loc[idx, "Altura"] = nova_altura
                    df_atual.loc[idx, "Data Atualização"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    salvar_dados(df_atual)
                    st.session_state["df_base"] = df_atual
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
            cod_int = st.text_input("Cód. Interno *").strip().upper()
            cod_fab = st.text_input("Cód. Fabricante *").strip().upper()
            desc = st.text_input("Descrição Completa *").strip().upper()
            rua = st.text_input("Rua *").strip().upper()
            box = st.text_input("Box *").strip().upper()
            altura = st.text_input("Altura *").strip().upper()
            
            btn_cad = st.form_submit_button("Cadastrar / Ocupar", use_container_width=True)
            if btn_cad:
                if not (cod_int and cod_fab and desc and rua and box and altura):
                    st.warning("Preencha todos os campos obrigatórios (*).")
                else:
                    df_atual = st.session_state["df_base"]
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
                    df_atual = pd.concat([df_atual, pd.DataFrame([novo_registro])], ignore_index=True)
                    salvar_dados(df_atual)
                    st.session_state["df_base"] = df_atual
                    st.success("✅ Novo produto cadastrado/endereçado com sucesso!")

# =========================================================
# TELA 4: LIMPAR ENDEREÇO
# =========================================================
elif opcao_menu == "🧹 Limpar Endereço":
    st.header("🧹 Desocupar / Limpar Endereço")
    if validar_admin():
        with st.form("form_limpar"):
            cod_limp = st.text_input("Código Interno ou Fabricante a Desocupar *").strip().upper()
            btn_limp = st.form_submit_button("Desocupar Endereço", use_container_width=True)
            
            if btn_limp:
                df_atual = st.session_state["df_base"]
                idx = df_atual[(df_atual["Cod. Interno"].str.upper() == cod_limp) | (df_atual["Cod. Fabricante"].str.upper() == cod_limp)].index
                if not idx.empty:
                    df_atual.loc[idx, ["Rua", "Box", "Altura"]] = ""
                    df_atual.loc[idx, "Data Atualização"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    salvar_dados(df_atual)
                    st.session_state["df_base"] = df_atual
                    st.success("✅ Endereço desocupado com sucesso!")
                else:
                    st.error("Produto não localizado no estoque.")

# =========================================================
# TELA 5: IMPORTAÇÃO / ATUALIZAÇÃO DA BASE EM MASSA (ADMIN)
# =========================================================
elif opcao_menu == "📥 Importar / Atualizar Base em Massa":
    st.header("📥 Importação e Atualização da Base em Massa")
    if validar_admin():
        st.write("Faça o upload do arquivo Excel (`.xlsx`) com o mapeamento completo do estoque para atualizar o WMS em tempo real.")
        
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
                    st.success("✅ Base de dados do WMS atualizada com sucesso! O novo mapeamento já está ativo para uso.")
            except Exception as e:
                st.error(f"Erro ao processar o arquivo Excel: {e}")

# =========================================================
# TELA 6: BACKUP E HISTÓRICO
# =========================================================
elif opcao_menu == "💾 Backup e Histórico":
    st.header("💾 Backup e Exportação da Base")
    st.write("Baixe uma cópia da base de estoque atualizada:")
    
    if os.path.exists(ARQUIVO_EXCEL):
        with open(ARQUIVO_EXCEL, "rb") as f:
            st.download_button(
                label="📥 Baixar Base_Estoque.xlsx Atualizada",
                data=f,
                file_name=f"Backup_WMS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
