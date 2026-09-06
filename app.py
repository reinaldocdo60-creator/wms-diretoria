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
# GERENCIAMENTO DE USUÁRIOS E SENHAS (COM PERSISTÊNCIA NO EXCEL)
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"

def carregar_usuarios():
    # Padrão inicial caso o arquivo ou aba não existam
    usuarios_padrao = {
        "operador": {"senha": "op123", "perfil": "OPERADOR"},
        "admin": {"senha": "admin123", "perfil": "ADMIN"}
    }
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            df_user = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Usuarios", dtype=str)
            df_user = df_user.fillna("")
            db = {}
            for _, row in df_user.iterrows():
                user = str(row["USUARIO"]).strip().lower()
                senha = str(row["SENHA"]).strip()
                perfil = str(row["PERFIL"]).strip().upper()
                if user:
                    db[user] = {"senha": senha, "perfil": perfil}
            return db if db else usuarios_padrao
        except Exception:
            return usuarios_padrao
    return usuarios_padrao

def salvar_usuarios(db):
    try:
        lista_user = []
        for user, info in db.items():
            lista_user.append({
                "USUARIO": user,
                "SENHA": info["senha"],
                "PERFIL": info["perfil"]
            })
        df_user = pd.DataFrame(lista_user)
        
        if os.path.exists(ARQUIVO_EXCEL):
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_user.to_excel(writer, sheet_name="Usuarios", index=False)
        else:
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
                df_user.to_excel(writer, sheet_name="Usuarios", index=False)
    except Exception as e:
        st.error(f"Erro ao salvar usuários: {e}")

if "usuarios_db" not in st.session_state:
    st.session_state["usuarios_db"] = carregar_usuarios()

# Inicialização da Sessão
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = ""
if "perfil" not in st.session_state:
    st.session_state["perfil"] = ""
if "linhas_congeladas" not in st.session_state:
    st.session_state["linhas_congeladas"] = pd.DataFrame()

if "q_busca" not in st.session_state:
    st.session_state["q_busca"] = ""
if "q_valid" not in st.session_state:
    st.session_state["q_valid"] = ""

# =========================================================
# CARREGAMENTO SEGURO DA BASE DE DADOS (BLINDADO PARA MOBILE)
# =========================================================
@st.cache_data(ttl=2)
def carregar_dados():
    colunas_padrao = ["CODINTERNO", "CODFAB", "DESCRICAO", "RUA", "BOX", "ALTURA", "GARANTIA", "CAIXA", "DATA ATUALIZACAO"]
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            try:
                df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str)
            except Exception:
                df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
            
            df = df.fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]
            return df
        except Exception:
            return pd.DataFrame(columns=colunas_padrao)
    else:
        return pd.DataFrame(columns=colunas_padrao)

def salvar_dados(df):
    db_atual = st.session_state["usuarios_db"]
    lista_user = [{"USUARIO": k, "SENHA": v["senha"], "PERFIL": v["perfil"]} for k, v in db_atual.items()]
    df_user = pd.DataFrame(lista_user)
    
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_Dados", index=False)
        df_user.to_excel(writer, sheet_name="Usuarios", index=False)
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

    usuarios_disponiveis = list(st.session_state["usuarios_db"].keys())
    usuario_input = st.selectbox("Selecione o Perfil / Usuário", usuarios_disponiveis)
    senha_input = st.text_input("Senha de Acesso", type="password")

    if st.button("🔑 Entrar no WMS", use_container_width=True):
        db = st.session_state["usuarios_db"]
        user_info = db.get(usuario_input.lower())
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
st.sidebar.title("📦 WMS LITLE")
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

if st.session_state["perfil"] == "ADMIN":
    opcoes_menu.append("👥 Gerenciar Usuários")

opcao_menu = st.sidebar.radio("Navegação Principal", opcoes_menu)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Sair do Sistema", use_container_width=True):
    st.session_state["autenticado"] = False
    st.session_state["usuario_logado"] = ""
    st.session_state["perfil"] = ""
    st.session_state["q_busca"] = ""
    st.session_state["q_valid"] = ""
    st.rerun()

def validar_admin():
    if st.session_state["perfil"] != "ADMIN":
        st.error("⚠️ Acesso restrito! Esta funcionalidade exige perfil de ADMINISTRADOR.")
        return False
    return True

# =========================================================
# TELA 1: PESQUISA E VALIDAÇÃO
# =========================================================
if opcao_menu == "🔍 Pesquisa e Validação (Geral)":
    st.header("🔍 Pesquisa e Validação")

    col_l1, col_l2 = st.columns([4, 1])
    with col_l2:
        if st.button("🧹 Limpar Busca", use_container_width=True):
            st.session_state["q_busca"] = ""
            st.session_state["q_valid"] = ""
            st.rerun()

    q_busca = st.text_input("1️⃣ PESQUISA (Cód., Descrição, Endereço):", key="q_busca").strip().upper()
    q_valid = st.text_input("2️⃣ VALIDAÇÃO (Bipe o Cód. Fabricante):", key="q_valid").strip().upper()

    df_res = st.session_state.get("df_base", carregar_dados()).copy()
    df_res.columns = [str(c).strip().upper() for c in df_res.columns]

    if q_busca and not df_res.empty:
        try:
            col_cod_int = "CODINTERNO" if "CODINTERNO" in df_res.columns else "COD. INTERNO"
            col_cod_fab = "CODFAB" if "CODFAB" in df_res.columns else "COD. FABRICANTE"
            col_desc = "DESCRICAO" if "DESCRICAO" in df_res.columns else "DESCRIÇÃO"
            
            mask = pd.Series(False, index=df_res.index)
            
            if col_cod_int in df_res.columns:
                mask = mask | df_res[col_cod_int].astype(str).str.upper().str.contains(q_busca, regex=False)
            if col_cod_fab in df_res.columns:
                mask = mask | df_res[col_cod_fab].astype(str).str.upper().str.contains(q_busca, regex=False)
            if col_desc in df_res.columns:
                mask = mask | df_res[col_desc].astype(str).str.upper().str.contains(q_busca, regex=False)
            if "RUA" in df_res.columns:
                mask = mask | df_res["RUA"].astype(str).str.upper().str.contains(q_busca, regex=False)
            if "BOX" in df_res.columns:
                mask = mask | df_res["BOX"].astype(str).str.upper().str.contains(q_busca, regex=False)
            if "ALTURA" in df_res.columns:
                mask = mask | df_res["ALTURA"].astype(str).str.upper().str.contains(q_busca, regex=False)
                
            df_res = df_res[mask]
        except Exception as e:
            st.error(f"Erro ao filtrar busca: {e}")
    elif not q_busca:
        df_res = df_res.iloc[0:0]

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
        col_validacao = "CODFAB" if "CODFAB" in df_res.columns else "COD. FABRICANTE"
        df_total = st.session_state.get("df_base", carregar_dados())
        df_total.columns = [str(c).strip().upper() for c in df_total.columns]
        
        if not df_total.empty and col_validacao in df_total.columns and (df_total[col_validacao].astype(str).str.upper() == q_valid).any():
            st.success(f"✅ VALIDAÇÃO OK: Código {q_valid} encontrado no estoque!")
        else:
            st.error(f"❌ ATENÇÃO: Código {q_valid} NÃO ENCONTRADO no estoque!")

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
                col_ci = "CODINTERNO" if "CODINTERNO" in df_atual.columns else "COD. INTERNO"
                col_cf = "CODFAB" if "CODFAB" in df_atual.columns else "COD. FABRICANTE"
                
                idx = df_atual[(df_atual[col_ci].str.upper() == cod_mover) | (df_atual[col_cf].str.upper() == cod_mover)].index
                if not idx.empty:
                    df_atual.loc[idx, "RUA"] = nova_rua
                    df_atual.loc[idx, "BOX"] = novo_box
                    df_atual.loc[idx, "ALTURA"] = nova_altura
                    if "DATA ATUALIZACAO" in df_atual.columns:
                        df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
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
                        "CODINTERNO": cod_int,
                        "CODFAB": cod_fab,
                        "DESCRICAO": desc,
                        "RUA": rua,
                        "BOX": box,
                        "ALTURA": altura,
                        "GARANTIA": "",
                        "CAIXA": "",
                        "DATA ATUALIZACAO": datetime.now().strftime("%Y-%m-%d %H:%M")
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
                col_ci = "CODINTERNO" if "CODINTERNO" in df_atual.columns else "COD. INTERNO"
                col_cf = "CODFAB" if "CODFAB" in df_atual.columns else "COD. FABRICANTE"
                
                idx = df_atual[(df_atual[col_ci].str.upper() == cod_limp) | (df_atual[col_cf].str.upper() == cod_limp)].index
                if not idx.empty:
                    df_atual.loc[idx, ["RUA", "BOX", "ALTURA"]] = ""
                    if "DATA ATUALIZACAO" in df_atual.columns:
                        df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
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
                df_novo.columns = [str(c).strip().upper() for c in df_novo.columns]
                
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

# =========================================================
# TELA 7: GERENCIAMENTO DE USUÁRIOS (ADMIN)
# =========================================================
elif opcao_menu == "👥 Gerenciar Usuários":
    st.header("👥 Gerenciamento de Colaboradores e Acessos")
    if validar_admin():
        db = st.session_state["usuarios_db"]
        
        dados_tabela = [{"Usuário": k.capitalize(), "Perfil": v["perfil"]} for k, v in db.items()]
        st.dataframe(pd.DataFrame(dados_tabela), use_container_width=True)
        
        st.divider()
        st.subheader("🔑 Alterar Senha de Usuário")
        
        col_alt1, col_alt2 = st.columns(2)
        with col_alt1:
            usuario_para_alterar = st.selectbox("Selecione o usuário para alterar a senha", list(db.keys()), key="sel_alt_user")
        with col_alt2:
            nova_senha_input = st.text_input("Nova Senha", type="password", key="txt_nova_senha")
            
        if st.button("💾 Salvar Nova Senha", use_container_width=True):
            if not nova_senha_input:
                st.warning("Digite a nova senha.")
            else:
                db[usuario_para_alterar]["senha"] = nova_senha_input
                salvar_usuarios(db)
                st.success(f"✅ Senha do usuário '{usuario_para_alterar.capitalize()}' alterada e salva com sucesso!")
        
        st.divider()
        st.subheader("Cadastrar Novo Colaborador")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            novo_usuario = st.text_input("Nome de Usuário").strip().lower()
        with col2:
            nova_senha = st.text_input("Senha", type="password")
        with col3:
            novo_perfil = st.selectbox("Perfil", ["OPERADOR", "ADMIN"])
        
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
                salvar_usuarios(st.session_state["usuarios_db"])
                st.success(f"Usuário '{novo_usuario.capitalize()}' cadastrado e salvo com sucesso!")
                st.rerun()
                
        st.divider()
        st.subheader("Remover Usuário")
        
        usuarios_removiveis = [u for u in db.keys() if u != "admin"]
        if usuarios_removiveis:
            usuario_para_remover = st.selectbox("Selecione o usuário para excluir", usuarios_removiveis)
            if st.button("Excluir Usuário Selecionado", type="primary"):
                if usuario_para_remover in st.session_state["usuarios_db"]:
                    del st.session_state["usuarios_db"][usuario_para_remover]
                    salvar_usuarios(st.session_state["usuarios_db"])
                    st.success(f"Usuário '{usuario_para_remover.capitalize()}' removido com sucesso!")
                    st.rerun()
        else:
            st.info("Não há outros usuários cadastrados para remoção.")
