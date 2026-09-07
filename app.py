import streamlit as str_lit
import pandas as pd
import os
import json
from datetime import datetime

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================
str_lit.set_page_config(
    page_title="WMS - Gestão de Estoque",
    page_icon="📦",
    layout="wide"
)

# =========================================================
# GERENCIAMENTO DE USUÁRIOS E SENHAS (PERSISTÊNCIA JSON)
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"
ARQUIVO_USUARIOS = "usuarios.json"

# Ordem oficial e unificada exigida em todo o sistema e importação
COLUNAS_PADRAO = [
    "GARANTIA", 
    "CODINTERNO", 
    "CODFAB", 
    "DESCRICAO", 
    "CAIXA", 
    "RUA", 
    "BOX", 
    "ALTURA", 
    "PALLET", 
    "PLT", 
    "DATA ATUALIZACAO"
]

def carregar_usuarios():
    usuarios_padrao = {
        "operador": {"senha": "op123", "perfil": "OPERADOR"},
        "admin": {"senha": "admin123", "perfil": "ADMIN"}
    }
    
    # Tenta carregar do JSON dedicado
    if os.path.exists(ARQUIVO_USUARIOS):
        try:
            with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as f:
                db = json.load(f)
                if db:
                    return db
        except Exception:
            pass
            
    # Fallback: Se não tem o JSON mas o Excel tem a aba Usuarios antiga, tenta resgatar de lá uma vez
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            xl = pd.ExcelFile(ARQUIVO_EXCEL)
            if "Usuarios" in xl.sheet_names:
                df_user = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Usuarios", dtype=str)
                df_user = df_user.fillna("")
                db = {}
                for _, row in df_user.iterrows():
                    user = str(row["USUARIO"]).strip().lower()
                    senha = str(row["SENHA"]).strip()
                    perfil = str(row["PERFIL"]).strip().upper()
                    if user:
                        db[user] = {"senha": senha, "perfil": perfil}
                if db:
                    salvar_usuarios(db) # Salva já no formato JSON novo
                    return db
        except Exception:
            pass
            
    return usuarios_padrao

def salvar_usuarios(db):
    try:
        with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=4)
        str_lit.session_state["usuarios_db"] = db
    except Exception as e:
        str_lit.error(f"Erro ao salvar usuários: {e}")

if "usuarios_db" not in str_lit.session_state:
    str_lit.session_state["usuarios_db"] = carregar_usuarios()

# Inicialização da Sessão
if "autenticado" not in str_lit.session_state:
    str_lit.session_state["autenticado"] = False
if "usuario_logado" not in str_lit.session_state:
    str_lit.session_state["usuario_logado"] = ""
if "perfil" not in str_lit.session_state:
    str_lit.session_state["perfil"] = ""
if "linhas_congeladas" not in str_lit.session_state:
    str_lit.session_state["linhas_congeladas"] = pd.DataFrame()

if "q_busca" not in str_lit.session_state:
    str_lit.session_state["q_busca"] = ""
if "q_valid" not in str_lit.session_state:
    str_lit.session_state["q_valid"] = ""

# =========================================================
# CARREGAMENTO SEGURO DA BASE DE DADOS (ESTOQUE)
# =========================================================
@str_lit.cache_data(ttl=2)
def carregar_dados():
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            xl = pd.ExcelFile(ARQUIVO_EXCEL)
            if "Base_Dados" in xl.sheet_names:
                df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str)
            else:
                df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
            
            df = df.fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            for col in COLUNAS_PADRAO:
                if col not in df.columns:
                    df[col] = ""
                    
            df = df[COLUNAS_PADRAO]
            return df
        except Exception:
            return pd.DataFrame(columns=COLUNAS_PADRAO)
    else:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

def salvar_dados(df):
    df = df.fillna("")
    for col in COLUNAS_PADRAO:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_PADRAO]
    
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_Dados", index=False)
    str_lit.cache_data.clear()

if "df_base" not in str_lit.session_state:
    str_lit.session_state["df_base"] = carregar_dados()

df_base = str_lit.session_state["df_base"]

# =========================================================
# TELA DE LOGIN
# =========================================================
if not str_lit.session_state["autenticado"]:
    str_lit.title("📦 WMS - Acesso ao Sistema")
    str_lit.subheader("🔒 Identificação do Usuário")

    str_lit.session_state["usuarios_db"] = carregar_usuarios()
    usuarios_disponiveis = list(str_lit.session_state["usuarios_db"].keys())
    
    usuario_input = str_lit.selectbox("Selecione o Perfil / Usuário", usuarios_disponiveis)
    senha_input = str_lit.text_input("Senha de Acesso", type="password")

    if str_lit.button("🔑 Entrar no WMS", use_container_width=True):
        db = str_lit.session_state["usuarios_db"]
        user_info = db.get(usuario_input.lower())
        if user_info and user_info["senha"] == senha_input:
            str_lit.session_state["autenticado"] = True
            str_lit.session_state["usuario_logado"] = usuario_input.capitalize()
            str_lit.session_state["perfil"] = user_info["perfil"]
            str_lit.rerun()
        else:
            str_lit.error("❌ Senha incorreta!")
    str_lit.stop()

# =========================================================
# BARRA LATERAL (MENU E PERFIL)
# =========================================================
str_lit.sidebar.title("📦 WMS LITLE")
str_lit.sidebar.write(f"👤 **Usuário:** {str_lit.session_state['usuario_logado']}")
str_lit.sidebar.write(f"🛡️ **Perfil:** `{str_lit.session_state['perfil']}`")

str_lit.sidebar.markdown("---")

opcoes_menu = [
    "🔍 Pesquisa e Validação (Geral)",
    "🚚 Mover Produto",
    "➕ Cadastrar / Ocupar",
    "🧹 Limpar Endereço",
    "📥 Importar / Atualizar Base em Massa",
    "💾 Backup e Histórico"
]

if str_lit.session_state["perfil"] == "ADMIN":
    opcoes_menu.append("👥 Gerenciar Usuários")

opcao_menu = str_lit.sidebar.radio("Navegação Principal", opcoes_menu)

str_lit.sidebar.markdown("---")
if str_lit.sidebar.button("🚪 Sair do Sistema", use_container_width=True):
    str_lit.session_state["autenticado"] = False
    str_lit.session_state["usuario_logado"] = ""
    str_lit.session_state["perfil"] = ""
    str_lit.session_state["q_busca"] = ""
    str_lit.session_state["q_valid"] = ""
    str_lit.rerun()

def validar_admin():
    if str_lit.session_state["perfil"] != "ADMIN":
        str_lit.error("⚠️ Acesso restrito! Esta funcionalidade exige perfil de ADMINISTRADOR.")
        return False
    return True

# =========================================================
# TELA 1: PESQUISA E VALIDAÇÃO
# =========================================================
if opcao_menu == "🔍 Pesquisa e Validação (Geral)":
    str_lit.header("🔍 Pesquisa e Validação")

    col_l1, col_l2 = str_lit.columns([4, 1])
    with col_l2:
        if str_lit.button("🧹 Limpar Busca", use_container_width=True):
            str_lit.session_state["q_busca"] = ""
            str_lit.session_state["q_valid"] = ""
            str_lit.rerun()

    q_busca = str_lit.text_input("1️⃣ PESQUISA (Cód., Descrição, Endereço):", key="q_busca").strip().upper()
    q_valid = str_lit.text_input("2️⃣ VALIDAÇÃO (Bipe o Cód. Fabricante):", key="q_valid").strip().upper()

    df_res = str_lit.session_state.get("df_base", carregar_dados()).copy()
    df_res.columns = [str(c).strip().upper() for c in df_res.columns]

    if q_busca and not df_res.empty:
        try:
            mask = pd.Series(False, index=df_res.index)
            for col in COLUNAS_PADRAO:
                if col in df_res.columns:
                    mask = mask | df_res[col].astype(str).str.upper().str.contains(q_busca, regex=False)
            df_res = df_res[mask]
        except Exception as e:
            str_lit.error(f"Erro ao filtrar busca: {e}")
    elif not q_busca:
        df_res = df_res.iloc[0:0]

    if str_lit.button("❄️ CONGELAR LINHAS DA PESQUISA", use_container_width=True):
        if not df_res.empty:
            congeladas_atuais = str_lit.session_state.get("linhas_congeladas", pd.DataFrame())
            str_lit.session_state["linhas_congeladas"] = pd.concat([congeladas_atuais, df_res]).drop_duplicates()
            str_lit.success("Linhas congeladas com sucesso!")

    if str_lit.button("🔥 LIMPAR LINHAS CONGELADAS", use_container_width=True):
        str_lit.session_state["linhas_congeladas"] = pd.DataFrame()
        str_lit.info("Acúmulo de linhas limpo!")

    tab1, tab2 = str_lit.tabs([f"🔎 Resultado ({len(df_res)})", f"❄️ Congeladas ({len(str_lit.session_state.get('linhas_congeladas', pd.DataFrame()))})"])

    with tab1:
        str_lit.dataframe(df_res, use_container_width=True)
        
        if str_lit.session_state["perfil"] == "ADMIN" and not df_res.empty:
            str_lit.markdown("---")
            str_lit.subheader("⚙️ Ação Rápida de Administrador (Desocupar Endereço)")
            str_lit.write("Selecione um dos produtos encontrados acima para limpar o endereço instantaneamente:")
            
            opcoes_desocupar = []
            mapa_opcoes = {}
            for _, r in df_res.iterrows():
                ci = str(r.get("CODINTERNO", ""))
                cf = str(r.get("CODFAB", ""))
                desc = str(r.get("DESCRICAO", ""))
                rua = str(r.get("RUA", ""))
                box = str(r.get("BOX", ""))
                rotulo = f"Cód. Int: {ci} | Cód. Fab: {cf} | {desc} (End: Rua {rua}, Box {box})"
                opcoes_desocupar.append(rotulo)
                mapa_opcoes[rotulo] = ci if ci else cf

            if opcoes_desocupar:
                col_sel_adm, col_btn_adm = str_lit.columns([3, 1])
                with col_sel_adm:
                    item_escolhido = str_lit.selectbox("Escolha o produto da lista acima para desocupar:", opcoes_desocupar, key="sel_desocupar_rapido")
                with col_btn_adm:
                    str_lit.write("") 
                    str_lit.write("")
                    if str_lit.button("🗑️ Desocupar Endereço", use_container_width=True, type="primary"):
                        codigo_alvo = mapa_opcoes[item_escolhido]
                        df_atual = str_lit.session_state["df_base"]
                        
                        idx = df_atual[(df_atual["CODINTERNO"].str.upper() == codigo_alvo.upper()) | (df_atual["CODFAB"].str.upper() == codigo_alvo.upper())].index
                        if not idx.empty:
                            df_atual.loc[idx, ["RUA", "BOX", "ALTURA", "PALLET", "PLT", "CAIXA"]] = ""
                            if "DATA ATUALIZACAO" in df_atual.columns:
                                df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                            salvar_dados(df_atual)
                            str_lit.session_state["df_base"] = df_atual
                            str_lit.success(f"✅ Endereço do produto '{codigo_alvo}' limpo com sucesso!")
                            str_lit.rerun()

    with tab2:
        str_lit.dataframe(str_lit.session_state.get("linhas_congeladas", pd.DataFrame()), use_container_width=True)

    if q_valid:
        df_total = str_lit.session_state.get("df_base", carregar_dados())
        df_total.columns = [str(c).strip().upper() for c in df_total.columns]
        
        if not df_total.empty and "CODFAB" in df_total.columns and (df_total["CODFAB"].astype(str).str.upper() == q_valid).any():
            str_lit.success(f"✅ VALIDAÇÃO OK: Código {q_valid} encontrado no estoque!")
        else:
            str_lit.error(f"❌ ATENÇÃO: Código {q_valid} NÃO ENCONTRADO no estoque!")

# =========================================================
# TELA 2: MOVER PRODUTO
# =========================================================
elif opcao_menu == "🚚 Mover Produto":
    str_lit.header("🚚 Movimentação Interna de Produto")
    
    with str_lit.form("form_mover"):
        cod_mover = str_lit.text_input("Código do Produto (Interno ou Fabricante) *").strip().upper()
        
        col_m1, col_m2 = str_lit.columns(2)
        with col_m1:
            nova_rua = str_lit.text_input("Nova Rua *").strip().upper()
            novo_box = str_lit.text_input("Novo Box *").strip().upper()
            nova_altura = str_lit.text_input("Nova Altura *").strip().upper()
        with col_m2:
            nova_caixa = str_lit.text_input("Nova Caixa").strip().upper()
            novo_pallet = str_lit.text_input("Novo Pallet").strip().upper()
            novo_plt = str_lit.text_input("Novo PLT").strip().upper()
        
        btn_mover = str_lit.form_submit_button("Confirmar Movimentação", use_container_width=True)
        
        if btn_mover:
            if not cod_mover or not nova_rua or not novo_box or not nova_altura:
                str_lit.warning("Preencha os campos obrigatórios (*).")
            else:
                df_atual = str_lit.session_state["df_base"]
                idx = df_atual[(df_atual["CODINTERNO"].str.upper() == cod_mover) | (df_atual["CODFAB"].str.upper() == cod_mover)].index
                if not idx.empty:
                    df_atual.loc[idx, "RUA"] = nova_rua
                    df_atual.loc[idx, "BOX"] = novo_box
                    df_atual.loc[idx, "ALTURA"] = nova_altura
                    if nova_caixa:
                        df_atual.loc[idx, "CAIXA"] = nova_caixa
                    if novo_pallet:
                        df_atual.loc[idx, "PALLET"] = novo_pallet
                    if novo_plt:
                        df_atual.loc[idx, "PLT"] = novo_plt
                        
                    if "DATA ATUALIZACAO" in df_atual.columns:
                        df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    salvar_dados(df_atual)
                    str_lit.session_state["df_base"] = df_atual
                    str_lit.success("✅ Produto movimentado com sucesso!")
                else:
                    str_lit.error("Produto não localizado no estoque.")

# =========================================================
# TELA 3: CADASTRAR / OCUPAR
# =========================================================
elif opcao_menu == "➕ Cadastrar / Ocupar":
    str_lit.header("➕ Cadastrar / Ocupar Endereço")
    if validar_admin():
        with str_lit.form("form_cadastrar"):
            col_c1, col_c2 = str_lit.columns(2)
            with col_c1:
                garantia = str_lit.text_input("1. GARANTIA").strip().upper()
                cod_int = str_lit.text_input("2. CODINTERNO (Cód. Interno) *").strip().upper()
                cod_fab = str_lit.text_input("3. CODFAB (Cód. Fabricante) *").strip().upper()
                desc = str_lit.text_input("4. DESCRICAO (Descrição Completa) *").strip().upper()
                caixa = str_lit.text_input("5. CAIXA").strip().upper()
            with col_c2:
                rua = str_lit.text_input("6. RUA *").strip().upper()
                box = str_lit.text_input("7. BOX *").strip().upper()
                altura = str_lit.text_input("8. ALTURA *").strip().upper()
                pallet = str_lit.text_input("9. PALLET").strip().upper()
                plt = str_lit.text_input("10. PLT").strip().upper()
            
            btn_cad = str_lit.form_submit_button("Cadastrar / Ocupar", use_container_width=True)
            if btn_cad:
                if not (cod_int and cod_fab and desc and rua and box and altura):
                    str_lit.warning("Preencha todos os campos obrigatórios (*).")
                else:
                    df_atual = str_lit.session_state["df_base"]
                    
                    novo_registro = {
                        "GARANTIA": garantia,
                        "CODINTERNO": cod_int,
                        "CODFAB": cod_fab,
                        "DESCRICAO": desc,
                        "CAIXA": caixa,
                        "RUA": rua,
                        "BOX": box,
                        "ALTURA": altura,
                        "PALLET": pallet,
                        "PLT": plt,
                        "DATA ATUALIZACAO": datetime.now().strftime("%Y-%m-%d %H:%M")
                    }
                    
                    df_novo_item = pd.DataFrame([novo_registro])[COLUNAS_PADRAO]
                    df_atual = pd.concat([df_atual, df_novo_item], ignore_index=True)
                    
                    salvar_dados(df_atual)
                    str_lit.session_state["df_base"] = df_atual
                    str_lit.success("✅ Novo produto cadastrado/endereçado com sucesso!")

# =========================================================
# TELA 4: LIMPAR ENDEREÇO
# =========================================================
elif opcao_menu == "🧹 Limpar Endereço":
    str_lit.header("🧹 Desocupar / Limpar Endereço")
    if validar_admin():
        with str_lit.form("form_limpar"):
            cod_limp = str_lit.text_input("Código Interno ou Fabricante a Desocupar *").strip().upper()
            btn_limp = str_lit.form_submit_button("Desocupar Endereço", use_container_width=True)
            
            if btn_limp:
                df_atual = str_lit.session_state["df_base"]
                idx = df_atual[(df_atual["CODINTERNO"].str.upper() == cod_limp) | (df_atual["CODFAB"].str.upper() == cod_limp)].index
                if not idx.empty:
                    df_atual.loc[idx, ["RUA", "BOX", "ALTURA", "PALLET", "PLT", "CAIXA"]] = ""
                    if "DATA ATUALIZACAO" in df_atual.columns:
                        df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    salvar_dados(df_atual)
                    str_lit.session_state["df_base"] = df_atual
                    str_lit.success("✅ Endereço desocupado com sucesso!")
                else:
                    str_lit.error("Produto não localizado no estoque.")

# =========================================================
# TELA 5: IMPORTAÇÃO / ATUALIZAÇÃO DA BASE EM MASSA
# =========================================================
elif opcao_menu == "📥 Importar / Atualizar Base em Massa":
    str_lit.header("📥 Importação e Atualização da Base em Massa")
    if validar_admin():
        str_lit.write("Faça o upload do arquivo Excel (`.xlsx`) com o mapeamento completo do estoque para atualizar o WMS em tempo real.")
        
        arquivo_enviado = str_lit.file_uploader("Selecione a planilha Excel mapeada", type=["xlsx"])
        
        if arquivo_enviado is not None:
            try:
                xl_up = pd.ExcelFile(arquivo_enviado)
                if "Base_Dados" in xl_up.sheet_names:
                    df_novo = pd.read_excel(arquivo_enviado, sheet_name="Base_Dados", dtype=str)
                else:
                    df_novo = pd.read_excel(arquivo_enviado, dtype=str)
                
                df_novo = df_novo.fillna("")
                df_novo.columns = [str(c).strip().upper() for c in df_novo.columns]
                
                for col in COLUNAS_PADRAO:
                    if col not in df_novo.columns:
                        df_novo[col] = ""
                df_novo = df_novo[COLUNAS_PADRAO]
                
                str_lit.success("Planilha lida com sucesso! Pré-visualização das 10 primeiras linhas:")
                str_lit.dataframe(df_novo.head(10), use_container_width=True)
                
                if str_lit.button("🚀 Confirmar e Atualizar Base do WMS", use_container_width=True):
                    salvar_dados(df_novo)
                    str_lit.session_state["df_base"] = df_novo
                    str_lit.balloons()
                    str_lit.success("✅ Base de dados do WMS atualizada com sucesso! O novo mapeamento já está ativo para uso.")
            except Exception as e:
                str_lit.error(f"Erro ao processar o arquivo Excel: {e}")

# =========================================================
# TELA 6: BACKUP E HISTÓRICO
# =========================================================
elif opcao_menu == "💾 Backup e Histórico":
    str_lit.header("💾 Backup e Exportação da Base")
    str_lit.write("Baixe uma cópia da base de estoque atualizada:")
    
    if os.path.exists(ARQUIVO_EXCEL):
        with open(ARQUIVO_EXCEL, "rb") as f:
            str_lit.download_button(
                label="📥 Baixar Base_Estoque.xlsx Atualizada",
                data=f,
                file_name=f"Backup_WMS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# =========================================================
# TELA 7: GERENCIAMENTO DE USUÁRIOS
# =========================================================
elif opcao_menu == "👥 Gerenciar Usuários":
    str_lit.header("👥 Gerenciamento de Colaboradores e Acessos")
    if validar_admin():
        str_lit.session_state["usuarios_db"] = carregar_usuarios()
        db = str_lit.session_state["usuarios_db"]
        
        dados_tabela = [{"Usuário": k.capitalize(), "Perfil": v["perfil"]} for k, v in db.items()]
        str_lit.dataframe(pd.DataFrame(dados_tabela), use_container_width=True)
        
        str_lit.divider()
        str_lit.subheader("🔑 Alterar Senha de Usuário")
        
        col_alt1, col_alt2 = str_lit.columns(2)
        with col_alt1:
            usuario_para_alterar = str_lit.selectbox("Selecione o usuário para alterar a senha", list(db.keys()), key="sel_alt_user")
        with col_alt2:
            nova_senha_input = str_lit.text_input("Nova Senha", type="password", key="txt_nova_senha")
            
        if str_lit.button("💾 Salvar Nova Senha", use_container_width=True):
            if not nova_senha_input:
                str_lit.warning("Digite a nova senha.")
            else:
                db[usuario_para_alterar]["senha"] = nova_senha_input
                salvar_usuarios(db)
                str_lit.success(f"✅ Senha do usuário '{usuario_para_alterar.capitalize()}' alterada e salva com sucesso!")
        
        str_lit.divider()
        str_lit.subheader("Cadastrar Novo Colaborador")
        
        col1, col2, col3 = str_lit.columns(3)
        with col1:
            novo_usuario = str_lit.text_input("Nome de Usuário").strip().lower()
        with col2:
            nova_senha = str_lit.text_input("Senha", type="password")
        with col3:
            novo_perfil = str_lit.selectbox("Perfil", ["OPERADOR", "ADMIN"])
        
        if str_lit.button("Cadastrar Usuário", use_container_width=True):
            if not novo_usuario or not nova_senha:
                str_lit.warning("Preencha o usuário e a senha.")
            elif novo_usuario in db:
                str_lit.error("Este usuário já existe!")
            else:
                db[novo_usuario] = {
                    "senha": nova_senha,
                    "perfil": novo_perfil
                }
                salvar_usuarios(db)
                str_lit.success(f"Usuário '{novo_usuario.capitalize()}' cadastrado e salvo com sucesso!")
                str_lit.rerun()
                
        str_lit.divider()
        str_lit.subheader("Remover Usuário")
        
        usuarios_removiveis = [u for u in db.keys() if u != "admin"]
        if usuarios_removiveis:
            usuario_para_remover = str_lit.selectbox("Selecione o usuário para excluir", usuarios_removiveis)
            if str_lit.button("Excluir Usuário Selecionado", type="primary"):
                if usuario_para_remover in db:
                    del db[usuario_para_remover]
                    salvar_usuarios(db)
                    str_lit.success(f"Usuário '{usuario_para_remover.capitalize()}' removido com sucesso!")
                    str_lit.rerun()
        else:
            str_lit.info("Não há outros usuários cadastrados para remoção.")
