import streamlit as st
import pandas as pd
import os
import time
import uuid
import hashlib
from contextlib import contextmanager
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
# CONSTANTES
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"
ARQUIVO_LOCK = ARQUIVO_EXCEL + ".lock"

# Campos de negócio (o que o usuário vê e preenche)
COLUNAS_NEGOCIO = [
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

# ID_ENDERECO é um identificador interno único por linha/endereço. Como o
# mesmo CODINTERNO pode existir em vários endereços ao mesmo tempo, esse ID
# é o que garante que "mover" ou "limpar" afete SÓ o endereço escolhido —
# nunca todas as linhas que compartilham o mesmo código.
COLUNAS_PADRAO = COLUNAS_NEGOCIO + ["ID_ENDERECO"]


# =========================================================
# TRAVA DE ARQUIVO (reduz risco de duas sessões gravando
# o Excel ao mesmo tempo e corrompendo o arquivo)
# =========================================================
@contextmanager
def _lock_arquivo(timeout=5):
    inicio = time.time()
    obtido = False
    while True:
        try:
            fd = os.open(ARQUIVO_LOCK, os.O_CREAT | os.O_EXCL | os.O_RDWR)
            os.close(fd)
            obtido = True
            break
        except FileExistsError:
            if time.time() - inicio > timeout:
                break
            time.sleep(0.1)
    try:
        yield obtido
    finally:
        if obtido:
            try:
                os.remove(ARQUIVO_LOCK)
            except Exception:
                pass


# =========================================================
# SENHAS: HASH (antes ficavam salvas em texto puro)
# =========================================================
def hash_senha(senha):
    return hashlib.sha256(str(senha).encode("utf-8")).hexdigest()


def _eh_hash_valido(valor):
    return isinstance(valor, str) and len(valor) == 64 and all(c in "0123456789abcdef" for c in valor.lower())


# =========================================================
# IDs ÚNICOS DE ENDEREÇO
# =========================================================
def _preencher_ids_faltantes(df):
    """Garante que toda linha tenha um ID_ENDERECO único (gera os que
    estiverem em branco — linhas antigas, recém-importadas, etc).
    Retorna (df_atualizado, precisou_gerar_algum: bool)."""
    df = df.copy()
    if "ID_ENDERECO" not in df.columns:
        df["ID_ENDERECO"] = ""
    df["ID_ENDERECO"] = df["ID_ENDERECO"].fillna("").astype(str)
    faltando = df["ID_ENDERECO"].str.strip() == ""
    precisou = bool(faltando.any())
    if precisou:
        df.loc[faltando, "ID_ENDERECO"] = [str(uuid.uuid4()) for _ in range(int(faltando.sum()))]
    return df, precisou


# =========================================================
# GERENCIAMENTO DE USUÁRIOS E SENHAS
# =========================================================
def carregar_usuarios():
    usuarios_padrao = {
        "operador": {"senha": hash_senha("op123"), "perfil": "OPERADOR"},
        "admin": {"senha": hash_senha("admin123"), "perfil": "ADMIN"}
    }
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
                return db if db else usuarios_padrao
        except KeyError:
            st.warning("A aba 'Usuarios' existe mas está com colunas inválidas. Usando usuários padrão.")
        except Exception:
            pass
    return usuarios_padrao


def _ler_base_bruta():
    """Leitura direta do disco, sem cache — usada nas operações de escrita
    para minimizar a janela de tempo entre ler e salvar os dados. Preenche
    IDs faltantes só em memória (não salva sozinha; quem chamar decide)."""
    if not os.path.exists(ARQUIVO_EXCEL):
        return pd.DataFrame(columns=COLUNAS_PADRAO)
    try:
        xl = pd.ExcelFile(ARQUIVO_EXCEL)
        if "Base_Dados" in xl.sheet_names:
            df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str)
        else:
            outras_abas = [s for s in xl.sheet_names if s != "Usuarios"]
            if outras_abas:
                df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=outras_abas[0], dtype=str)
            else:
                return pd.DataFrame(columns=COLUNAS_PADRAO)
    except Exception:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

    df = df.fillna("")
    df.columns = [str(c).strip().upper() for c in df.columns]
    for col in COLUNAS_PADRAO:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_PADRAO]
    df, _ = _preencher_ids_faltantes(df)
    return df


def salvar_usuarios(db):
    try:
        df_base_atual = _ler_base_bruta()

        lista_user = []
        for user, info in db.items():
            lista_user.append({
                "USUARIO": user,
                "SENHA": info["senha"],
                "PERFIL": info["perfil"]
            })
        df_user = pd.DataFrame(lista_user)

        with _lock_arquivo() as obtido:
            if not obtido:
                st.error("Não foi possível salvar: o arquivo está sendo usado por outra operação. Tente novamente.")
                return
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
                df_base_atual.to_excel(writer, sheet_name="Base_Dados", index=False)
                df_user.to_excel(writer, sheet_name="Usuarios", index=False)

        st.session_state["usuarios_db"] = carregar_usuarios()
        carregar_dados.clear()
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
# CARREGAMENTO SEGURO DA BASE DE DADOS (versão em cache,
# usada apenas para EXIBIÇÃO. Escritas usam _ler_base_bruta()
# para pegar sempre o dado mais recente do disco).
# =========================================================
@st.cache_data(ttl=2)
def carregar_dados():
    return _ler_base_bruta()


def salvar_dados(df):
    """Salva a base de produtos. Retorna True em caso de sucesso."""
    db_atual = st.session_state.get("usuarios_db", carregar_usuarios())
    lista_user = [{"USUARIO": k, "SENHA": v["senha"], "PERFIL": v["perfil"]} for k, v in db_atual.items()]
    df_user = pd.DataFrame(lista_user)

    df, _ = _preencher_ids_faltantes(df)
    df = df.fillna("")
    for col in COLUNAS_PADRAO:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_PADRAO]

    with _lock_arquivo() as obtido:
        if not obtido:
            st.error("Não foi possível salvar: o arquivo está sendo usado por outra operação. Tente novamente em instantes.")
            return False
        with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Base_Dados", index=False)
            df_user.to_excel(writer, sheet_name="Usuarios", index=False)

    carregar_dados.clear()
    return True


if "df_base" not in st.session_state:
    df_inicial = carregar_dados()
    # Migração única por sessão: se a base ainda não tem ID_ENDERECO
    # preenchido em todas as linhas (base antiga, importação anterior),
    # gera e persiste agora.
    df_inicial, precisou_migrar = _preencher_ids_faltantes(df_inicial)
    if precisou_migrar:
        salvar_dados(df_inicial)
        carregar_dados.clear()
        df_inicial = carregar_dados()
    st.session_state["df_base"] = df_inicial

df_base = st.session_state["df_base"]


def _rotulo_endereco(row):
    """Monta um texto legível para o usuário identificar um endereço
    específico quando o mesmo Código Interno aparece em vários lugares."""
    return (
        f"Rua {row.get('RUA', '') or '—'} | Box {row.get('BOX', '') or '—'} | "
        f"Caixa {row.get('CAIXA', '') or '—'} | Altura {row.get('ALTURA', '') or '—'} "
        f"— {row.get('DESCRICAO', '')}"
    )


# =========================================================
# TELA DE LOGIN
# =========================================================
if not st.session_state["autenticado"]:
    st.title("📦 WMS - Acesso ao Sistema")
    st.subheader("🔒 Identificação do Usuário")

    st.session_state["usuarios_db"] = carregar_usuarios()
    usuarios_disponiveis = list(st.session_state["usuarios_db"].keys())

    usuario_input = st.selectbox("Selecione o Perfil / Usuário", usuarios_disponiveis)
    senha_input = st.text_input("Senha de Acesso", type="password")

    if st.button("🔑 Entrar no WMS", use_container_width=True):
        db = st.session_state["usuarios_db"]
        user_info = db.get(usuario_input.lower())
        autenticado_ok = False

        if user_info:
            senha_armazenada = user_info["senha"]
            if _eh_hash_valido(senha_armazenada):
                autenticado_ok = (hash_senha(senha_input) == senha_armazenada)
            else:
                # Compatibilidade com bases antigas que ainda guardam a senha
                # em texto puro: aceita uma vez e converte para hash na hora.
                if senha_armazenada == senha_input:
                    autenticado_ok = True
                    db[usuario_input.lower()]["senha"] = hash_senha(senha_input)
                    salvar_usuarios(db)

        if autenticado_ok:
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
            mask = pd.Series(False, index=df_res.index)
            # Busca só nos campos de negócio — ID_ENDERECO é um UUID interno
            # e não faz sentido pesquisável pelo usuário.
            for col in COLUNAS_NEGOCIO:
                if col in df_res.columns:
                    mask = mask | df_res[col].astype(str).str.upper().str.contains(q_busca, regex=False)
            df_res = df_res[mask]
        except Exception as e:
            st.error(f"Erro ao filtrar busca: {e}")
    elif not q_busca:
        df_res = df_res.iloc[0:0]

    if st.button("❄️ CONGELAR LINHAS DA PESQUISA", use_container_width=True):
        if not df_res.empty:
            congeladas_atuais = st.session_state.get("linhas_congeladas", pd.DataFrame())
            st.session_state["linhas_congeladas"] = pd.concat([congeladas_atuais, df_res]).drop_duplicates(subset=["ID_ENDERECO"])
            st.success("Linhas congeladas com sucesso!")

    if st.button("🔥 LIMPAR LINHAS CONGELADAS", use_container_width=True):
        st.session_state["linhas_congeladas"] = pd.DataFrame()
        st.info("Acúmulo de linhas limpo!")

    tab1, tab2 = st.tabs([f"🔎 Resultado ({len(df_res)})", f"❄️ Congeladas ({len(st.session_state.get('linhas_congeladas', pd.DataFrame()))})"])

    with tab1:
        st.dataframe(df_res.drop(columns=["ID_ENDERECO"], errors="ignore"), use_container_width=True)

        if st.session_state["perfil"] == "ADMIN" and not df_res.empty:
            st.markdown("---")
            st.subheader("⚙️ Ação Rápida de Administrador (Desocupar Endereço)")
            st.write("Selecione um dos produtos encontrados acima para limpar o endereço instantaneamente:")

            opcoes_desocupar = []
            mapa_opcoes = {}
            for _, r in df_res.iterrows():
                ci = str(r.get("CODINTERNO", ""))
                cf = str(r.get("CODFAB", ""))
                rotulo = f"Cód. Int: {ci} | Cód. Fab: {cf} | {_rotulo_endereco(r)}"
                opcoes_desocupar.append(rotulo)
                # Mapeia pelo ID_ENDERECO único da linha, não pelo código —
                # assim a ação afeta só ESTE endereço, mesmo que o mesmo
                # código exista em outros lugares.
                mapa_opcoes[rotulo] = r["ID_ENDERECO"]

            if opcoes_desocupar:
                col_sel_adm, col_btn_adm = st.columns([3, 1])
                with col_sel_adm:
                    item_escolhido = st.selectbox("Escolha o endereço da lista acima para desocupar:", opcoes_desocupar, key="sel_desocupar_rapido")
                with col_btn_adm:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Desocupar Endereço", use_container_width=True, type="primary"):
                        id_alvo = mapa_opcoes[item_escolhido]
                        # Lê a base direto do disco antes de alterar, para não
                        # sobrescrever mudanças feitas por outra sessão/usuário
                        # enquanto esta ficava aberta.
                        df_atual = _ler_base_bruta()

                        idx = df_atual[df_atual["ID_ENDERECO"] == id_alvo].index
                        if not idx.empty:
                            df_atual.loc[idx, ["RUA", "BOX", "ALTURA", "PALLET", "PLT", "CAIXA"]] = ""
                            if "DATA ATUALIZACAO" in df_atual.columns:
                                df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                            if salvar_dados(df_atual):
                                st.session_state["df_base"] = df_atual
                                st.success("✅ Endereço limpo com sucesso!")
                                st.rerun()
                        else:
                            st.error("Esse endereço não existe mais na base (pode ter sido alterado por outra pessoa). Atualize a página e tente novamente.")

    with tab2:
        st.dataframe(st.session_state.get("linhas_congeladas", pd.DataFrame()).drop(columns=["ID_ENDERECO"], errors="ignore"), use_container_width=True)

    if q_valid:
        df_total = st.session_state.get("df_base", carregar_dados())
        df_total.columns = [str(c).strip().upper() for c in df_total.columns]

        if not df_total.empty and "CODFAB" in df_total.columns and (df_total["CODFAB"].astype(str).str.upper() == q_valid).any():
            qtd_enderecos = (df_total["CODFAB"].astype(str).str.upper() == q_valid).sum()
            extra = f" (presente em {qtd_enderecos} endereços)" if qtd_enderecos > 1 else ""
            st.success(f"✅ VALIDAÇÃO OK: Código {q_valid} encontrado no estoque!{extra}")
        else:
            st.error(f"❌ ATENÇÃO: Código {q_valid} NÃO ENCONTRADO no estoque!")

# =========================================================
# TELA 2: MOVER PRODUTO
# =========================================================
elif opcao_menu == "🚚 Mover Produto":
    st.header("🚚 Movimentação Interna de Produto")

    cod_mover = st.text_input("Código do Produto (Interno ou Fabricante) *", key="cod_mover_busca").strip().upper()

    id_alvo = None
    if cod_mover:
        df_busca = _ler_base_bruta()
        candidatos = df_busca[(df_busca["CODINTERNO"].str.upper() == cod_mover) | (df_busca["CODFAB"].str.upper() == cod_mover)]

        if candidatos.empty:
            st.error("Produto não localizado no estoque.")
        elif len(candidatos) == 1:
            id_alvo = candidatos.iloc[0]["ID_ENDERECO"]
            st.info(f"Endereço atual: {_rotulo_endereco(candidatos.iloc[0])}")
        else:
            st.warning(f"Esse código está presente em {len(candidatos)} endereços diferentes. Selecione qual deseja mover:")
            opcoes = {_rotulo_endereco(r): r["ID_ENDERECO"] for _, r in candidatos.iterrows()}
            rotulo_escolhido = st.selectbox("Endereço a mover:", list(opcoes.keys()), key="sel_mover_endereco")
            id_alvo = opcoes[rotulo_escolhido]

    with st.form("form_mover"):
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            nova_rua = st.text_input("Nova Rua *").strip().upper()
            novo_box = st.text_input("Novo Box *").strip().upper()
            nova_altura = st.text_input("Nova Altura *").strip().upper()
        with col_m2:
            nova_caixa = st.text_input("Nova Caixa").strip().upper()
            novo_pallet = st.text_input("Novo Pallet").strip().upper()
            novo_plt = st.text_input("Novo PLT").strip().upper()

        btn_mover = st.form_submit_button("Confirmar Movimentação", use_container_width=True)

        if btn_mover:
            if not cod_mover or not nova_rua or not novo_box or not nova_altura:
                st.warning("Preencha os campos obrigatórios (*) e informe um código válido acima.")
            elif id_alvo is None:
                st.error("Nenhum endereço válido selecionado para mover.")
            else:
                df_atual = _ler_base_bruta()
                idx = df_atual[df_atual["ID_ENDERECO"] == id_alvo].index
                if idx.empty:
                    st.error("Esse endereço não existe mais na base (pode ter sido alterado por outra pessoa). Refaça a busca.")
                else:
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
                    if salvar_dados(df_atual):
                        st.session_state["df_base"] = df_atual
                        st.success("✅ Produto movimentado com sucesso!")

# =========================================================
# TELA 3: CADASTRAR / OCUPAR (NA ORDEM EXATA)
# =========================================================
elif opcao_menu == "➕ Cadastrar / Ocupar":
    st.header("➕ Cadastrar / Ocupar Endereço")
    if validar_admin():
        st.caption("O mesmo Código Interno pode ser cadastrado em vários endereços diferentes. "
                   "Só não é permitido repetir exatamente o mesmo produto no mesmo endereço (Rua + Box).")
        with st.form("form_cadastrar"):
            # Ordem exata: GARANTIA, CODINTERNO, CODFAB, DESCRICAO, CAIXA, RUA, BOX, ALTURA, PALLET, PLT
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                garantia = st.text_input("1. GARANTIA").strip().upper()
                cod_int = st.text_input("2. CODINTERNO (Cód. Interno) *").strip().upper()
                cod_fab = st.text_input("3. CODFAB (Cód. Fabricante) *").strip().upper()
                desc = st.text_input("4. DESCRICAO (Descrição Completa) *").strip().upper()
                caixa = st.text_input("5. CAIXA").strip().upper()
            with col_c2:
                rua = st.text_input("6. RUA *").strip().upper()
                box = st.text_input("7. BOX *").strip().upper()
                altura = st.text_input("8. ALTURA *").strip().upper()
                pallet = st.text_input("9. PALLET").strip().upper()
                plt = st.text_input("10. PLT").strip().upper()

            btn_cad = st.form_submit_button("Cadastrar / Ocupar", use_container_width=True)
            if btn_cad:
                if not (cod_int and cod_fab and desc and rua and box and altura):
                    st.warning("Preencha todos os campos obrigatórios (*).")
                else:
                    df_atual = _ler_base_bruta()

                    # Só bloqueia se for EXATAMENTE o mesmo produto no mesmo
                    # endereço (mesmo Código Interno + mesma Rua + mesmo Box).
                    # Múltiplos endereços para o mesmo produto são permitidos.
                    duplicado = (
                        (df_atual["CODINTERNO"].str.upper() == cod_int)
                        & (df_atual["RUA"].str.upper() == rua)
                        & (df_atual["BOX"].str.upper() == box)
                    ).any()

                    if duplicado:
                        st.error(f"❌ Já existe o produto '{cod_int}' cadastrado exatamente na Rua {rua}, Box {box}!")
                    else:
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
                            "DATA ATUALIZACAO": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "ID_ENDERECO": str(uuid.uuid4()),
                        }

                        df_novo_item = pd.DataFrame([novo_registro])[COLUNAS_PADRAO]
                        df_atual = pd.concat([df_atual, df_novo_item], ignore_index=True)

                        if salvar_dados(df_atual):
                            st.session_state["df_base"] = df_atual
                            st.success("✅ Novo produto cadastrado/endereçado com sucesso!")

# =========================================================
# TELA 4: LIMPAR ENDEREÇO
# =========================================================
elif opcao_menu == "🧹 Limpar Endereço":
    st.header("🧹 Desocupar / Limpar Endereço")
    if validar_admin():
        cod_limp = st.text_input("Código Interno ou Fabricante a Desocupar *", key="cod_limp_busca").strip().upper()

        id_alvo_limp = None
        if cod_limp:
            df_busca = _ler_base_bruta()
            candidatos = df_busca[(df_busca["CODINTERNO"].str.upper() == cod_limp) | (df_busca["CODFAB"].str.upper() == cod_limp)]

            if candidatos.empty:
                st.error("Produto não localizado no estoque.")
            elif len(candidatos) == 1:
                id_alvo_limp = candidatos.iloc[0]["ID_ENDERECO"]
                st.info(f"Endereço a limpar: {_rotulo_endereco(candidatos.iloc[0])}")
            else:
                st.warning(f"Esse código está presente em {len(candidatos)} endereços diferentes. Selecione qual deseja limpar:")
                opcoes_limp = {_rotulo_endereco(r): r["ID_ENDERECO"] for _, r in candidatos.iterrows()}
                rotulo_escolhido_limp = st.selectbox("Endereço a limpar:", list(opcoes_limp.keys()), key="sel_limpar_endereco")
                id_alvo_limp = opcoes_limp[rotulo_escolhido_limp]

        with st.form("form_limpar"):
            btn_limp = st.form_submit_button("Desocupar Endereço", use_container_width=True)

            if btn_limp:
                if not cod_limp:
                    st.warning("Informe o código do produto acima.")
                elif id_alvo_limp is None:
                    st.error("Nenhum endereço válido selecionado para limpar.")
                else:
                    df_atual = _ler_base_bruta()
                    idx = df_atual[df_atual["ID_ENDERECO"] == id_alvo_limp].index
                    if idx.empty:
                        st.error("Esse endereço não existe mais na base (pode ter sido alterado por outra pessoa). Refaça a busca.")
                    else:
                        df_atual.loc[idx, ["RUA", "BOX", "ALTURA", "PALLET", "PLT", "CAIXA"]] = ""
                        if "DATA ATUALIZACAO" in df_atual.columns:
                            df_atual.loc[idx, "DATA ATUALIZACAO"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                        if salvar_dados(df_atual):
                            st.session_state["df_base"] = df_atual
                            st.success("✅ Endereço desocupado com sucesso!")

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

                # Gera um ID_ENDERECO único para cada linha importada
                df_novo, _ = _preencher_ids_faltantes(df_novo)

                st.success("Planilha lida com sucesso! Pré-visualização das 10 primeiras linhas:")
                st.dataframe(df_novo.drop(columns=["ID_ENDERECO"], errors="ignore").head(10), use_container_width=True)

                st.warning("⚠️ Esta ação SUBSTITUI toda a base de produtos atual pelo conteúdo da planilha importada.")
                if st.button("🚀 Confirmar e Atualizar Base do WMS", use_container_width=True):
                    if salvar_dados(df_novo):
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

    # Restrito a ADMIN: o arquivo contém a aba "Usuarios" com credenciais
    # de acesso. Antes, qualquer usuário logado podia baixar esse arquivo
    # e obter a lista de usuários/senhas do sistema.
    if validar_admin():
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
        else:
            st.info("Ainda não há uma base salva para baixar.")

# =========================================================
# TELA 7: GERENCIAMENTO DE USUÁRIOS (ADMIN)
# =========================================================
elif opcao_menu == "👥 Gerenciar Usuários":
    st.header("👥 Gerenciamento de Colaboradores e Acessos")
    if validar_admin():
        st.session_state["usuarios_db"] = carregar_usuarios()
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
                db[usuario_para_alterar]["senha"] = hash_senha(nova_senha_input)
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
                db[novo_usuario] = {
                    "senha": hash_senha(nova_senha),
                    "perfil": novo_perfil
                }
                salvar_usuarios(db)
                st.success(f"Usuário '{novo_usuario.capitalize()}' cadastrado e salvo com sucesso!")
                st.rerun()

        st.divider()
        st.subheader("Remover Usuário")

        usuario_atual_lower = st.session_state["usuario_logado"].lower()
        usuarios_removiveis = [u for u in db.keys() if u != "admin" and u != usuario_atual_lower]
        if usuarios_removiveis:
            usuario_para_remover = st.selectbox("Selecione o usuário para excluir", usuarios_removiveis)
            if st.button("Excluir Usuário Selecionado", type="primary"):
                if usuario_para_remover in db:
                    del db[usuario_para_remover]
                    salvar_usuarios(db)
                    st.success(f"Usuário '{usuario_para_remover.capitalize()}' removido com sucesso!")
                    st.rerun()
        else:
            st.info("Não há outros usuários cadastrados para remoção.")
