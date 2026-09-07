import streamlit as str_lit
import pandas as pd
import os
import json
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

# Configuração da chave para o Gemini
try:
    if "GEMINI_API_KEY" in str_lit.secrets:
        genai.configure(api_key=str_lit.secrets["GEMINI_API_KEY"])
    else:
        genai.configure(api_key="AQ.Ab8RN6I1bXpNqsD3J7zN046MfeH4ld3DmrwraC1srEuEAnczaA")
except Exception:
    pass

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================
str_lit.set_page_config(
    page_title="WMS - Gestão de Estoque",
    page_icon="📦",
    layout="wide"
)

# =========================================================
# ESTILO CSS GLOBAL + REGRAS DE IMPRESSÃO (MEDIA PRINT)
# =========================================================
str_lit.markdown("""
    <style>
    @media print {
        /* Esconde a barra lateral (sidebar) */
        [data-testid="stSidebar"] {
            display: none !important;
        }
        /* Esconde o cabeçalho superior do Streamlit */
        header {
            display: none !important;
        }
        /* Esconde botões, caixas de texto e inputs de filtro na hora da impressão */
        .stTextInput, .stButton, .stTabs, div[data-testid="column"] {
            display: none !important;
        }
        /* Garante que o corpo principal ocupe a página toda sem margens indesejadas */
        .main {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

# =========================================================
# GERENCIAMENTO DE USUÁRIOS E SENHAS (PERSISTÊNCIA JSON)
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"
ARQUIVO_USUARIOS = "usuarios.json"

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
    
    if os.path.exists(ARQUIVO_USUARIOS):
        try:
            with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as f:
                db = json.load(f)
                if db:
                    return db
        except Exception:
            pass
            
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
                    salvar_usuarios(db)
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

if "autenticado" not in str_lit.session_state:
    str_lit.session_state["autenticado"] = False
if "usuario_logado" not in str_lit.session_state:
    str_lit.session_state["usuario_logado"] = ""
if "perfil" not in str_lit.session_state:
    str_lit.session_state["perfil"] = ""
if "linhas_congeladas" not in str_lit.session_state:
    str_lit.session_state["linhas_congeladas"] = pd.DataFrame()
if "modo_impressao" not in str_lit.session_state:
    str_lit.session_state["modo_impressao"] = False
if "modo_impressao_inventario" not in str_lit.session_state:
    str_lit.session_state["modo_impressao_inventario"] = False
if "modo_impressao_manual" not in str_lit.session_state:
    str_lit.session_state["modo_impressao_manual"] = False

# Controle de tentativas de login incorretas e bloqueio temporário
if "tentativas_login" not in str_lit.session_state:
    str_lit.session_state["tentativas_login"] = 0
if "tempo_bloqueio" not in str_lit.session_state:
    str_lit.session_state["tempo_bloqueio"] = None

if "q_busca" not in str_lit.session_state:
    str_lit.session_state["q_busca"] = ""
if "q_valid" not in str_lit.session_state:
    str_lit.session_state["q_valid"] = ""

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
            
            # Limpa espaços e converte para maiúsculo
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # Mapeamento de sinônimos para evitar colunas em branco
            mapa_colunas = {
                "PALET": "PALLET",
                "PALETE": "PALLET",
                "COD.INTERNO": "CODINTERNO",
                "COD_INTERNO": "CODINTERNO",
                "CÓDIGO INTERNO": "CODINTERNO",
                "COD.FAB": "CODFAB",
                "COD_FAB": "CODFAB",
                "FABRICANTE": "CODFAB",
                "DESC": "DESCRICAO",
                "DESCRIÇÃO": "DESCRICAO",
                "ULTIMA ATUALIZACAO": "DATA ATUALIZACAO",
                "DATA_ATUALIZACAO": "DATA ATUALIZACAO"
            }
            df = df.rename(columns=mapa_colunas)
            
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
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Base_Dados"
    
    ws.views.sheetView[0].showGridLines = True
    
    cor_cabecalho_fundo = "1F4E78"
    cor_cabecalho_fonte = "FFFFFF"
    cor_linha_alternada = "F2F5F9"
    
    fonte_cabecalho = Font(name="Arial", size=11, bold=True, color=cor_cabecalho_fonte)
    preenchimento_cabecalho = PatternFill(start_color=cor_cabecalho_fundo, end_color=cor_cabecalho_fundo, fill_type="solid")
    
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    alinhar_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alinhar_esquerda = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    ws.row_dimensions[1].height = 28
    for col_num, col_name in enumerate(COLUNAS_PADRAO, 1):
        cell = ws.cell(row=1, column=col_num, value=col_name)
        cell.font = fonte_cabecalho
        cell.fill = preenchimento_cabecalho
        cell.alignment = alinhar_centro
        cell.border = borda_fina
        
    for row_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        ws.row_dimensions[row_idx].height = 20
        is_par = (row_idx % 2 == 0)
        fill_atual = PatternFill(start_color=cor_linha_alternada, end_color=cor_linha_alternada, fill_type="solid") if is_par else None
        
        for col_idx, valor in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=str(valor))
            cell.font = Font(name="Arial", size=10)
            cell.border = borda_fina
            
            if fill_atual:
                cell.fill = fill_atual
                
            nome_coluna = COLUNAS_PADRAO[col_idx - 1]
            if nome_coluna == "DESCRICAO":
                cell.alignment = alinhar_esquerda
            else:
                cell.alignment = alinhar_centro

    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)
        
    wb.save(ARQUIVO_EXCEL)
    str_lit.cache_data.clear()

if "df_base" not in str_lit.session_state:
    str_lit.session_state["df_base"] = carregar_dados()

df_base = str_lit.session_state["df_base"]

# =========================================================
# TELA DE LOGIN
# =========================================================
if not str_lit.session_state["autenticado"]:
    _, col_centro, _ = str_lit.columns([1, 1.2, 1])
    
    with col_centro:
        str_lit.markdown("<br>", unsafe_allow_html=True)
        str_lit.title("📦 WMS - Acesso")
        str_lit.subheader("🔒 Identificação")

        bloqueado_ate = str_lit.session_state.get("tempo_bloqueio", None)
        
        if bloqueado_ate and datetime.now() < bloqueado_ate:
            str_lit.error("❌ Muitas tentativas incorretas. Acesso bloqueado por 1 minuto.")
            if str_lit.button("🔄 Verificar se o tempo acabou", use_container_width=True):
                str_lit.rerun()
            str_lit.stop()
        elif bloqueado_ate and datetime.now() >= bloqueado_ate:
            str_lit.session_state["tempo_bloqueio"] = None
            str_lit.session_state["tentativas_login"] = 0

        str_lit.session_state["usuarios_db"] = carregar_usuarios()
        usuarios_disponiveis = list(str_lit.session_state["usuarios_db"].keys())
        
        usuario_input = str_lit.selectbox("Selecione o Perfil / Usuário", usuarios_disponiveis)
        senha_input = str_lit.text_input("Senha de Acesso", type="password")

        tentativas_restantes = 3 - str_lit.session_state["tentativas_login"]
        str_lit.info(f"⚠️ Tentativas restantes: **{tentativas_restantes}**")

        if str_lit.button("🔑 Entrar no WMS", use_container_width=True):
            db = str_lit.session_state["usuarios_db"]
            user_info = db.get(usuario_input.lower())
            
            if user_info and user_info["senha"] == senha_input:
                str_lit.session_state["autenticado"] = True
                str_lit.session_state["usuario_logado"] = usuario_input.capitalize()
                str_lit.session_state["perfil"] = user_info["perfil"]
                str_lit.session_state["tentativas_login"] = 0
                str_lit.session_state["tempo_bloqueio"] = None
                str_lit.rerun()
            else:
                str_lit.session_state["tentativas_login"] += 1
                if str_lit.session_state["tentativas_login"] >= 3:
                    str_lit.session_state["tempo_bloqueio"] = datetime.now() + timedelta(minutes=1)
                    str_lit.error("❌ Limite de 3 tentativas atingido. Bloqueado por 1 minuto.")
                    str_lit.rerun()
                else:
                    str_lit.error(f"❌ Senha incorreta! Tentativa {str_lit.session_state['tentativas_login']} de 3.")
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
    "🧹 Limpar Endereço / Excluir Linha",
    "📥 Importar / Atualizar Base em Massa",
    "📊 Sugestão de Inventário Mensal",
    "💾 Backup e Histórico",
    "📖 Manual de Instruções",
    "🤖 Assistente IA"
]

if str_lit.session_state["perfil"] == "ADMIN":
    opcoes_menu.insert(6, "👥 Gerenciar Usuários")

opcao_menu = str_lit.sidebar.radio("Navegação Principal", opcoes_menu)

str_lit.sidebar.markdown("---")
if str_lit.sidebar.button("🚪 Sair do Sistema", use_container_width=True):
    str_lit.session_state["autenticado"] = False
    str_lit.session_state["usuario_logado"] = ""
    str_lit.session_state["perfil"] = ""
    str_lit.session_state["q_busca"] = ""
    str_lit.session_state["q_valid"] = ""
    str_lit.session_state["modo_impressao"] = False
    str_lit.session_state["modo_impressao_inventario"] = False
    str_lit.session_state["modo_impressao_manual"] = False
    str_lit.rerun()

def validar_admin():
    if str_lit.session_state["perfil"] != "ADMIN":
        str_lit.error("⚠️ Acesso restrito! Esta funcionalidade exige perfil de ADMINISTRADOR.")
        return False
    return True

# =========================================================
# TELA 1: PESQUISA E VALIDAÇÃO (COM FILTRO DE GARANTIA)
# =========================================================
if opcao_menu == "🔍 Pesquisa e Validação (Geral)":
    
    df_res = str_lit.session_state.get("df_base", carregar_dados()).copy()
    df_res.columns = [str(c).strip().upper() for c in df_res.columns]

    if "GARANTIA" not in df_res.columns:
        df_res["GARANTIA"] = ""

    # Extrai as garantias únicas presentes na base para o seletor
    garantias_disponiveis = ["TODAS"]
    unicas = sorted([str(x).strip().upper() for x in df_res["GARANTIA"].unique() if str(x).strip() != ""])
    garantias_disponiveis.extend(unicas)

    if str_lit.session_state.get("modo_impressao", False):
        str_lit.markdown("## 🖨️ Pré-visualização de Impressão - WMS")
        str_lit.info("Esta é a visualização limpa pronta para impressão. Pressione **Ctrl + P** no seu teclado para enviar diretamente à impressora ou salvar em PDF.")
        
        if str_lit.button("⬅️ Voltar para a Pesquisa Normal", use_container_width=True):
            str_lit.session_state["modo_impressao"] = False
            str_lit.rerun()
            
        str_lit.markdown("---")
        
        q_busca_ativo = str_lit.session_state.get("q_busca", "").strip().upper()
        df_res_print = df_res.copy()
        
        if q_busca_ativo and not df_res_print.empty:
            mask = pd.Series(False, index=df_res_print.index)
            for col in COLUNAS_PADRAO:
                if col in df_res_print.columns:
                    mask = mask | df_res_print[col].astype(str).str.upper().str.contains(q_busca_ativo, regex=False)
            df_res_print = df_res_print[mask]
        else:
            df_res_print = df_res_print.iloc[0:0]
            
        str_lit.write(f"**Filtro aplicado:** {q_busca_ativo if q_busca_ativo else 'Nenhum'} | **Total de registros:** {len(df_res_print)}")
        str_lit.table(df_res_print)
        str_lit.stop()

    str_lit.header("🔍 Pesquisa e Validação")

    col_l1, col_l2 = str_lit.columns([4, 1])
    with col_l2:
        str_lit.write("")
        str_lit.write("")
        if str_lit.button("🧹 Limpar Busca", use_container_width=True):
            str_lit.session_state["q_busca"] = ""
            str_lit.session_state["q_valid"] = ""
            str_lit.rerun()

    # Seletor de Garantia em destaque logo acima da busca
    garantia_selecionada = str_lit.selectbox(
        "📌 SELECIONE A GARANTIA (Filtro Anti-Duplicidade):", 
        garantias_disponiveis,
        key="filtro_garantia_selectbox"
    )

    q_busca = str_lit.text_input("1️⃣ PESQUISA (Cód., Descrição, Endereço):", key="q_busca").strip().upper()
    q_valid = str_lit.text_input("2️⃣ VALIDAÇÃO (Bipe o Cód. Fabricante ou Interno da peça separada):", key="q_valid").strip().upper()

    # Aplicação do Filtro de Garantia primeiro
    if garantia_selecionada != "TODAS":
        df_res = df_res[df_res["GARANTIA"].astype(str).str.upper() == garantia_selecionada]

    # Aplicação do Filtro de Texto em seguida
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
        col_cab, col_btn_imp = str_lit.columns([3, 1])
        with col_cab:
            str_lit.markdown("### Resultado da Pesquisa de Estoque")
        with col_btn_imp:
            if str_lit.button("🖨️ Visualizar / Imprimir", use_container_width=True):
                str_lit.session_state["modo_impressao"] = True
                str_lit.rerun()

        str_lit.table(df_res)

    with tab2:
        str_lit.markdown("### Linhas Congeladas")
        str_lit.table(str_lit.session_state.get("linhas_congeladas", pd.DataFrame()))

    # VALIDAÇÃO BASEADA NO RETORNO DA PESQUISA ATUAL
    if q_valid:
        if not q_busca:
            str_lit.warning("⚠️ Para validar, faça primeiro uma pesquisa para trazer o item esperado na tela.")
        elif df_res.empty:
            str_lit.error("❌ VALIDAÇÃO FALHOU: Nenhum item encontrado na pesquisa atual para validar.")
        else:
            match_interno = (df_res["CODINTERNO"].astype(str).str.upper() == q_valid).any()
            match_fab = (df_res["CODFAB"].astype(str).str.upper() == q_valid).any()
            
            if match_interno or match_fab:
                str_lit.success(f"✅ VALIDAÇÃO OK! O código '{q_valid}' confere com o item pesquisado/separado.")
            else:
                str_lit.error(f"❌ ATENÇÃO: O código '{q_valid}' NÃO CONFERE com os itens listados na pesquisa atual! Verifique se a peça está correta.")

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
# TELA 4: LIMPAR ENDEREÇO / EXCLUIR LINHA
# =========================================================
elif opcao_menu == "🧹 Limpar Endereço / Excluir Linha":
    str_lit.header("🧹 Gerenciamento de Exclusão de Endereços/Linhas")
    if validar_admin():
        str_lit.write("Digite o código (Interno ou Fabricante) para localizar todas as ocorrências e linhas associadas a ele:")
        cod_busca_limpeza = str_lit.text_input("Código para consulta de exclusão *").strip().upper()
        
        if cod_busca_limpeza:
            df_atual = str_lit.session_state["df_base"]
            df_encontrados = df_atual[
                (df_atual["CODINTERNO"].str.upper() == cod_busca_limpeza) | 
                (df_atual["CODFAB"].str.upper() == cod_busca_limpeza)
            ]
            total_ocorrencias = len(df_encontrados)
            
            if total_ocorrencias > 0:
                str_lit.warning(f"⚠️ Atenção! Foram encontradas **{total_ocorrencias}** ocorrência(s) (endereço(s)) para o código **{cod_busca_limpeza}**.")
                
                df_exibicao = df_encontrados.copy()
                df_exibicao.insert(0, "LINHA_EXCEL", [idx + 2 for idx in df_encontrados.index])
                str_lit.table(df_exibicao)
                
                str_lit.markdown("---")
                str_lit.subheader("🗑️ Opções de Exclusão")
                
                opcoes_linhas = []
                mapa_linhas = {}
                for idx_real, r in df_encontrados.iterrows():
                    linha_excel = idx_real + 2
                    rua = r.get("RUA", "-")
                    box = r.get("BOX", "-")
                    texto_opcao = f"Linha {linha_excel} da Planilha (Rua: {rua}, Box: {box})"
                    opcoes_linhas.append(texto_opcao)
                    mapa_linhas[texto_opcao] = idx_real
                
                col_ex1, col_ex2 = str_lit.columns(2)
                with col_ex1:
                    str_lit.write("**Excluir apenas uma linha específica:**")
                    linha_escolhida = str_lit.selectbox("Selecione qual linha deseja excluir:", opcoes_linhas, key="sel_linha_excluir")
                    if str_lit.button("🔥 Excluir Linha Selecionada", use_container_width=True, type="primary"):
                        idx_alvo = mapa_linhas[linha_escolhida]
                        df_atual = df_atual.drop(idx_alvo).reset_index(drop=True)
                        salvar_dados(df_atual)
                        str_lit.session_state["df_base"] = df_atual
                        str_lit.success("✅ Linha selecionada excluída com sucesso!")
                        str_lit.rerun()
                        
                with col_ex2:
                    str_lit.write("**Excluir TODAS as ocorrências deste código:**")
                    str_lit.write("") 
                    if str_lit.button(f"🚨 Excluir TODAS as {total_ocorrencias} linhas deste código", use_container_width=True, type="primary"):
                        indices_alvo = df_encontrados.index
                        df_atual = df_atual.drop(indices_alvo).reset_index(drop=True)
                        salvar_dados(df_atual)
                        str_lit.session_state["df_base"] = df_atual
                        str_lit.success(f"✅ Todas as {total_ocorrencias} ocorrências do código '{cod_busca_limpeza}' foram excluídas com sucesso!")
                        str_lit.rerun()
            else:
                str_lit.info(f"Nenhum registro encontrado com o código '{cod_busca_limpeza}'.")

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
                
                mapa_colunas = {
                    "PALET": "PALLET",
                    "PALETE": "PALLET",
                    "COD.INTERNO": "CODINTERNO",
                    "COD_INTERNO": "CODINTERNO",
                    "CÓDIGO INTERNO": "CODINTERNO",
                    "COD.FAB": "CODFAB",
                    "COD_FAB": "CODFAB",
                    "FABRICANTE": "CODFAB",
                    "DESC": "DESCRICAO",
                    "DESCRIÇÃO": "DESCRICAO",
                    "ULTIMA ATUALIZACAO": "DATA ATUALIZACAO",
                    "DATA_ATUALIZACAO": "DATA ATUALIZACAO"
                }
                df_novo = df_novo.rename(columns=mapa_colunas)
                
                for col in COLUNAS_PADRAO:
                    if col not in df_novo.columns:
                        df_novo[col] = ""
                df_novo = df_novo[COLUNAS_PADRAO]
                
                str_lit.success("Planilha lida com sucesso! Pré-visualização das 10 primeiras linhas:")
                str_lit.table(df_novo.head(10))
                
                if str_lit.button("🚀 Confirmar e Atualizar Base do WMS", use_container_width=True):
                    salvar_dados(df_novo)
                    str_lit.session_state["df_base"] = df_novo
                    str_lit.balloons()
                    str_lit.success("✅ Base de dados do WMS atualizada com sucesso! O novo mapeamento já está ativo para uso.")
            except Exception as e:
                str_lit.error(f"Erro ao processar o arquivo Excel: {e}")

# =========================================================
# TELA 6: SUGESTÃO DE INVENTÁRIO MENSAL
# =========================================================
elif opcao_menu == "📊 Sugestão de Inventário Mensal":
    
    if str_lit.session_state.get("modo_impressao_inventario", False):
        str_lit.markdown("## 🖨️ Pré-visualização de Impressão - Sugestão de Inventário")
        str_lit.info("Esta é a visualização limpa pronta para impressão. Pressione **Ctrl + P** no seu teclado para enviar diretamente à impressora ou salvar em PDF.")
        
        if str_lit.button("⬅️ Voltar para a Sugestão Normal", use_container_width=True):
            str_lit.session_state["modo_impressao_inventario"] = False
            str_lit.rerun()
            
        str_lit.markdown("---")
        
        df_amostra_print = str_lit.session_state.get("df_base", carregar_dados()).copy()
        if not df_amostra_print.empty:
            df_amostra_print = df_amostra_print.head(10)
        else:
            df_amostra_print = df_amostra_print.iloc[0:0]
            
        str_lit.write(f"**Relatório de Amostragem para Inventário Cíclico** | **Total de itens:** {len(df_amostra_print)}")
        str_lit.table(df_amostra_print)
        str_lit.stop()

    str_lit.header("📊 Sugestão de Inventário por Amostragem")
    str_lit.info("💡 Esta é uma listagem de sugestão direcionada para o inventário cíclico mensal. Focada em itens prioritários da base para conferência física preventiva.")
    
    df_inv = str_lit.session_state.get("df_base", carregar_dados()).copy()
    if not df_inv.empty:
        df_inv = df_inv.head(20) # Sugere os 20 primeiros itens
    
    col_inv_cab, col_inv_btn = str_lit.columns([3, 1])
    with col_inv_cab:
        str_lit.markdown("### Amostragem Sugerida para Auditoria")
    with col_inv_btn:
        if str_lit.button("🖨️ Imprimir Inventário", use_container_width=True):
            str_lit.session_state["modo_impressao_inventario"] = True
            str_lit.rerun()

    str_lit.table(df_inv)

# =========================================================
# TELA: BACKUP E HISTÓRICO (EXEMPLO DE TELA ADICIONAL)
# =========================================================
elif opcao_menu == "💾 Backup e Histórico":
    str_lit.header("💾 Backup e Histórico da Base")
    str_lit.write("Faça o download do backup da base de dados atual em formato Excel.")
    if os.path.exists(ARQUIVO_EXCEL):
        with open(ARQUIVO_EXCEL, "rb") as f:
            str_lit.download_button(
                label="📥 Baixar Backup do Excel (.xlsx)",
                data=f,
                file_name=f"Backup_WMS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        str_lit.warning("Nenhum arquivo de base de dados encontrado para backup.")

# =========================================================
# TELA: MANUAL DE INSTRUÇÕES
# =========================================================
elif opcao_menu == "📖 Manual de Instruções":
    str_lit.header("📖 Manual de Instruções do WMS")
    str_lit.markdown("""
    Bem-vindo ao manual de operações do **WMS Litle**:
    - **Pesquisa e Validação:** Utilize para consultar endereços, caixas e validar peças bipe a bipe.
    - **Mover Produto:** Altere a rua, box ou altura de um produto mantendo o histórico atualizado.
    - **Cadastrar / Ocupar:** Adicione novos itens ou novos endereços à base (Perfil Admin).
    - **Limpar Endereço:** Remova ocorrências ou limpe registros obsoletos da base.
    - **Assistente IA:** Tire dúvidas sobre processos e rotinas diretamente com a Inteligência Artificial.
    """)

# =========================================================
# TELA: GERENCIAR USUÁRIOS (APENAS ADMIN)
# =========================================================
elif opcao_menu == "👥 Gerenciar Usuários":
    str_lit.header("👥 Gerenciamento de Usuários do Sistema")
    if validar_admin():
        db_users = str_lit.session_state.get("usuarios_db", {})
        
        # Exibe usuários cadastrados
        df_users_display = pd.DataFrame([
            {"USUARIO": u, "PERFIL": info["perfil"]} for u, info in db_users.items()
        ])
        str_lit.subheader("Usuários Ativos")
        str_lit.table(df_users_display)
        
        str_lit.markdown("---")
        str_lit.subheader("Adicionar Novo Usuário")
        with str_lit.form("form_novo_usuario"):
            novo_user = str_lit.text_input("Nome de Usuário").strip().lower()
            nova_senha = str_lit.text_input("Senha", type="password").strip()
            novo_perfil = str_lit.selectbox("Perfil", ["OPERADOR", "ADMIN"])
            btn_add_user = str_lit.form_submit_button("Cadastrar Usuário", use_container_width=True)
            
            if btn_add_user:
                if not novo_user or not nova_senha:
                    str_lit.warning("Preencha o usuário e a senha.")
                else:
                    db_users[novo_user] = {"senha": nova_senha, "perfil": novo_perfil}
                    salvar_usuarios(db_users)
                    str_lit.success(f"✅ Usuário '{novo_user}' cadastrado com sucesso!")
                    str_lit.rerun()
# =========================================================
# TELA: ASSISTENTE VIRTUAL DE IA (GEMINI)
# =========================================================
elif opcao_menu == "🤖 Assistente IA":
    str_lit.title("🤖 Assistente Virtual WMS")
    str_lit.markdown("Tire suas dúvidas sobre as rotinas, processos e regras de negócio do nosso sistema de gerenciamento de armazém.")

    if "historico_chat" not in str_lit.session_state:
        str_lit.session_state["historico_chat"] = []

    for mensagem in str_lit.session_state["historico_chat"]:
        with str_lit.chat_message(mensagem["role"]):
            str_lit.markdown(mensagem["content"])

    if duvida_usuario := str_lit.chat_input("Digite sua dúvida sobre o WMS ou operações de armazém..."):
        str_lit.session_state["historico_chat"].append({"role": "user", "content": duvida_usuario})
        with str_lit.chat_message("user"):
            str_lit.markdown(duvida_usuario)

        with str_lit.chat_message("assistant"):
            with str_lit.spinner("Consultando as regras do WMS..."):
                try:
                    from google import genai
                    
                    # Cole aqui a sua chave completa que começa com AQ.:
                    api_key_valor = "AQ.Ab8RN6LnwMvVl9Q3cYVjAm2A173bLltqeSYaqn5QK_EF8ERojg"
                    
                    # Inicializa o cliente oficial moderno
                    client = genai.Client(api_key=api_key_valor)
                    
                    prompt_sistema = "Você é o assistente virtual oficial de um Sistema de Gestão de Armazém (WMS). Responda dúvidas sobre as rotinas, processos e regras de negócio do sistema de forma clara, prestativa e em português brasileiro."
                    
                    # Usando o modelo padrão recomendado para essa arquitetura
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=f"{prompt_sistema}\n\nDúvida do usuário: {duvida_usuario}",
                    )
                    
                    resposta_ia = response.text
                    str_lit.markdown(resposta_ia)
                    str_lit.session_state["historico_chat"].append({"role": "assistant", "content": resposta_ia})

                except Exception as erro:
                    str_lit.error(f"Desculpe, ocorreu um erro ao consultar a IA: {erro}")
