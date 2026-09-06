import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="WMS - Sistema de Estoque",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO_EXCEL = "Base_Estoque.xlsx"
TEXTO_DISPONIVEL = "DISPONÍVEL"

# CSS Personalizado para Impressão e Alertas de Validação
st.markdown("""
    <style>
    .main { background-color: #F8FAFC; }
    
    .st-validado {
        background-color: #D1FAE5;
        color: #065F46;
        padding: 14px;
        border-radius: 8px;
        font-weight: bold;
        font-size: 18px;
        text-align: center;
        border: 2px solid #10B981;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    
    .st-nao-validado {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 14px;
        border-radius: 8px;
        font-weight: bold;
        font-size: 18px;
        text-align: center;
        border: 2px solid #EF4444;
        margin-top: 10px;
        margin-bottom: 15px;
    }

    @media print {
        [data-testid="stSidebar"], header, .stButton, .no-print, .validacao-container { display: none !important; }
        .main { background-color: white !important; }
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# CARREGAMENTO E MANIPULAÇÃO DE DADOS
# ============================================================
def carregar_dados():
    if not os.path.exists(ARQUIVO_EXCEL):
        df_vazio = pd.DataFrame(columns=[
            "CURVA", "GARANTIA", "CODINTERNO", "CODFAB", 
            "DESCRICAO", "CAIXA", "RUA", "BOX", "ALTURA", "PALLETE", "PLT"
        ])
        df_historico = pd.DataFrame(columns=["DATA_HORA", "USUARIO", "ACAO", "COD_INTERNO", "DESCRICAO", "ORIGEM", "DESTINO"])
        
        with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
            df_vazio.to_excel(writer, sheet_name="Base_Dados", index=False)
            df_historico.to_excel(writer, sheet_name="Historico", index=False)
            
    df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")
    return df_base

def salvar_dados(df_base):
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_base.to_excel(writer, sheet_name="Base_Dados", index=False)

def registrar_historico(acao, cod_int, desc, origem, destino, usuario):
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    novo_log = pd.DataFrame([{
        "DATA_HORA": data_hora, "USUARIO": usuario, "ACAO": acao,
        "COD_INTERNO": cod_int, "DESCRICAO": desc, "ORIGEM": origem, "DESTINO": destino
    }])
    
    try:
        df_hist = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Historico", dtype=str).fillna("")
    except Exception:
        df_hist = pd.DataFrame(columns=["DATA_HORA", "USUARIO", "ACAO", "COD_INTERNO", "DESCRICAO", "ORIGEM", "DESTINO"])
        
    df_hist = pd.concat([df_hist, novo_log], ignore_index=True)
    
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_hist.to_excel(writer, sheet_name="Historico", index=False)

# ============================================================
# ESTADO DA SESSÃO (LOGIN E LISTA CONGELADA / ACUMULADA)
# ============================================================
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "perfil_logado" not in st.session_state:
    st.session_state.perfil_logado = None
if "lista_congelada" not in st.session_state:
    st.session_state.lista_congelada = pd.DataFrame()

# Tela de Login
if st.session_state.usuario_logado is None:
    st.title("📦 WMS - Acesso ao Sistema")
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.subheader("🔐 Identificação do Usuário")
        usuario = st.selectbox("Selecione o Perfil / Usuário", ["operador", "admin"])
        senha = st.text_input("Senha de Acesso", type="password")
        
        if st.button("🔑 Entrar no WMS", use_container_width=True, type="primary"):
            if usuario == "admin" and senha == "admin123":
                st.session_state.usuario_logado = "Administrador"
                st.session_state.perfil_logado = "admin"
                st.rerun()
            elif usuario == "operador" and senha == "op123":
                st.session_state.usuario_logado = "Operador"
                st.session_state.perfil_logado = "operador"
                st.rerun()
            else:
                st.error("❌ Senha incorreta!")
    st.stop()

# ============================================================
# MENU LATERAL
# ============================================================
st.sidebar.title("📦 WMS Nuvem")
st.sidebar.markdown(f"👤 **Usuário:** `{st.session_state.usuario_logado}`")
st.sidebar.markdown(f"🛡️ **Perfil:** `{st.session_state.perfil_logado.upper()}`")

opcao_menu = st.sidebar.radio(
    "Navegação Principal",
    [
        "🔍 Pesquisa e Validação (Geral)", 
        "🚚 Mover Produto", 
        "➕ Cadastrar / Ocupar", 
        "🧹 Limpar Endereço", 
        "💾 Backup e Histórico"
    ]
)

if st.sidebar.button("🚪 Sair do Sistema"):
    st.session_state.usuario_logado = None
    st.session_state.perfil_logado = None
    st.session_state.lista_congelada = pd.DataFrame()
    st.rerun()

df_base = carregar_dados()

def validar_admin():
    if st.session_state.perfil_logado != "admin":
        st.error("⚠️ Acesso Negado: Esta funcionalidade exige privilégios de Administrador.")
        return False
    return True

# ============================================================
# TELA 1: CONSULTA, VALIDAÇÃO E CONGELAMENTO EM SEQUÊNCIA
# ============================================================
if opcao_menu == "🔍 Pesquisa e Validação (Geral)":
    st.header("🔍 Pesquisa, Validação e Congelamento de Endereços")
    
    # ------------------------------------------------------------
    # BLOCO DE BARRAS EM SEQUÊNCIA (PESQUISA + VALIDAÇÃO)
    # ------------------------------------------------------------
    st.subheader("📌 Painel de Operação")
    
    # 1. Primeira Barra: Pesquisa Geral
    busca = st.text_input("1️⃣ PESQUISA: Digite Cód. Interno, Fabricante, Descrição, Rua, Box ou Altura:").strip().upper()
    
    resultado_busca = df_base.copy()
    if busca:
        filtro = (
            resultado_busca["CODINTERNO"].str.upper().str.contains(busca, na=False) |
            resultado_busca["CODFAB"].str.upper().str.contains(busca, na=False) |
            resultado_busca["DESCRICAO"].str.upper().str.contains(busca, na=False) |
            resultado_busca["RUA"].str.upper().str.contains(busca, na=False) |
            resultado_busca["BOX"].str.upper().str.contains(busca, na=False) |
            resultado_busca["ALTURA"].str.upper().str.contains(busca, na=False)
        )
        resultado_busca = resultado_busca[filtro]
    else:
        resultado_busca = pd.DataFrame()

    # 2. Segunda Barra: Validação por Bipagem (logo em sequência)
    cod_bipado = st.text_input("2️⃣ VALIDAÇÃO: Bipe ou digite o Cód. Fabricante para conferência:", key="input_bipagem_seq").strip().upper()

    # Processamento da Validação com Áudios Profissionais
    if cod_bipado:
        base_comparacao = st.session_state.lista_congelada if not st.session_state.lista_congelada.empty else resultado_busca
        
        if base_comparacao.empty:
            st.warning("⚠️ Faça uma pesquisa ou congele linhas antes de bipar para conferência!")
        else:
            codigos_esperados = [str(c).strip().upper() for c in base_comparacao["CODFAB"].tolist()]
            
            if cod_bipado in codigos_esperados:
                item_validado = base_comparacao[base_comparacao["CODFAB"].str.strip().str.upper() == cod_bipado].iloc[0]
                
                # Som Profissional: Beep de Confirmação Scanner (Frequência Alta / Limpo)
                st.components.v1.html("""
                    <audio autoplay>
                        <source src="https://cdn.freesound.org/previews/351/351566_6142149-lq.mp3" type="audio/mpeg">
                    </audio>
                """, height=0)

                st.markdown(f"""
                    <div class="st-validado">
                        ✅ VALIDADO! <br>
                        Código: <b>{item_validado['CODFAB']}</b> | Produto: <b>{item_validado['DESCRICAO']}</b><br>
                        Localização: Rua {item_validado['RUA']} | Box {item_validado['BOX']} | Altura {item_validado['ALTURA']}
                    </div>
                """, unsafe_allow_html=True)
                
                registrar_historico("VALIDACAO_SUCESSO", item_validado['CODINTERNO'], item_validado['DESCRICAO'], cod_bipado, "VALIDADO_OK", st.session_state.usuario_logado)
            else:
                # Som Profissional: Alerta de Divergência/Erro Industrial
                st.components.v1.html("""
                    <audio autoplay>
                        <source src="https://cdn.freesound.org/previews/142/142608_1840739-lq.mp3" type="audio/mpeg">
                    </audio>
                """, height=0)

                st.markdown(f"""
                    <div class="st-nao-validado">
                        ❌ NÃO VALIDADO! <br>
                        O código <b>{cod_bipado}</b> NÃO confere com a pesquisa/congelamento atual!<br>
                        ⚠️ Item divergente. Evite a separação incorreta!
                    </div>
                """, unsafe_allow_html=True)
                
                registrar_historico("VALIDACAO_ERRO", "-", "DIVERGENTE", cod_bipado, "NAO_VALIDADO", st.session_state.usuario_logado)

    # ------------------------------------------------------------
    # BOTÕES DE AÇÃO DO CONGELAMENTO / IMPRESSÃO
    # ------------------------------------------------------------
    st.markdown("---")
    c_btn1, c_btn2, c_btn3 = st.columns(3)
    
    if c_btn1.button("❄️ CONGELAR LINHAS DA PESQUISA", use_container_width=True, type="primary"):
        if not resultado_busca.empty:
            st.session_state.lista_congelada = pd.concat([st.session_state.lista_congelada, resultado_busca]).drop_duplicates()
            st.success(f"✅ {len(resultado_busca)} linha(s) congelada(s) e adicionada(s) ao lote de impressão!")
        else:
            st.warning("⚠️ Realize uma pesquisa válida antes de congelar.")

    if c_btn2.button("🔥 DESCONGELAR / LIMPAR ACÚMULO", use_container_width=True):
        st.session_state.lista_congelada = pd.DataFrame()
        st.info("🔥 Lista congelada e acumulada foi limpa.")
        st.rerun()
        
    if c_btn3.button("🖨️ IMPRIMIR ACÚMULO (Ctrl + P)", use_container_width=True):
        if not st.session_state.lista_congelada.empty:
            st.info("💡 Dica: Pressione **Ctrl + P** no seu navegador para imprimir as linhas congeladas abaixo.")
        else:
            st.warning("⚠️ Nenhuma linha congelada para imprimir.")

    # ------------------------------------------------------------
    # TABELAS DE EXIBIÇÃO DE RESULTADOS
    # ------------------------------------------------------------
    tab_busca, tab_congelado = st.tabs([
        f"🔎 Resultado da Busca ({len(resultado_busca)})", 
        f"❄️ Linhas Acumuladas / Congeladas ({len(st.session_state.lista_congelada)})"
    ])
    
    with tab_busca:
        st.dataframe(resultado_busca, use_container_width=True, hide_index=True)
        
    with tab_congelado:
        if st.session_state.lista_congelada.empty:
            st.write("Nenhuma linha congelada no momento.")
        else:
            st.markdown(f"""
                <div style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 5px; margin-bottom: 15px;">
                    <h3>LISTA DE CONFERÊNCIA E ENDEREÇOS PARA IMPRESSÃO</h3>
                    <p>Emissão: {datetime.now().strftime('%d/%m/%Y - %H:%M')} | Operador: {st.session_state.usuario_logado}</p>
                </div>
            """, unsafe_allow_html=True)
            st.dataframe(st.session_state.lista_congelada, use_container_width=True, hide_index=True)

# ============================================================
# TELA 2: MOVER PRODUTO
# ============================================================
elif opcao_menu == "🚚 Mover Produto":
    st.header("🚚 Movimentação Física de Produto")
    
    if validar_admin():
        cod_fabricante = st.text_input("🔑 Digite o Código do Fabricante para localizar o item:").strip().upper()
        
        if cod_fabricante:
            mascara = df_base["CODFAB"].str.strip().str.upper() == cod_fabricante
            idx = df_base.index[mascara].tolist()
            
            if idx:
                pos = idx[0]
                cod_int = df_base.at[pos, 'CODINTERNO']
                desc = df_base.at[pos, 'DESCRICAO']
                rua_origem = df_base.at[pos, 'RUA']
                box_origem = df_base.at[pos, 'BOX']
                alt_origem = df_base.at[pos, 'ALTURA']
                plt_origem = df_base.at[pos, 'PALLETE']
                num_plt_origem = df_base.at[pos, 'PLT']
                
                origem_str = f"Rua:{rua_origem} Box:{box_origem} Alt:{alt_origem}"
                
                st.info(f"📌 **Produto Encontrado:** {cod_int} - {desc}")
                st.warning(f"📍 **Localização Atual:** {origem_str} | Pallet: {plt_origem} | PLT: {num_plt_origem}")
                
                with st.form("form_movimentacao"):
                    st.subheader("🎯 Informar Novo Endereço de Destino")
                    col_m1, col_m2, col_m3 = st.columns(3)
                    n_rua = col_m1.text_input("Nova Rua *", rua_origem).strip().upper()
                    n_box = col_m2.text_input("Novo Box *", box_origem).strip().upper()
                    n_alt = col_m3.text_input("Nova Altura *", alt_origem).strip().upper()
                    
                    col_m4, col_m5 = st.columns(2)
                    n_plt = col_m4.text_input("Novo Pallet", plt_origem).strip().upper()
                    n_plt_num = col_m5.text_input("Novo PLT", num_plt_origem).strip().upper()
                    
                    btn_mover = st.form_submit_button("💾 CONFIRMAR MOVIMENTAÇÃO", type="primary")
                    
                    if btn_mover:
                        if not (n_rua and n_box and n_alt):
                            st.error("⚠️ Os campos Rua, Box e Altura são obrigatórios!")
                        else:
                            destino_str = f"Rua:{n_rua} Box:{n_box} Alt:{n_alt}"
                            
                            df_base.at[pos, 'CODINTERNO'] = TEXTO_DISPONIVEL
                            df_base.at[pos, 'CODFAB'] = ""
                            df_base.at[pos, 'DESCRICAO'] = TEXTO_DISPONIVEL
                            df_base.at[pos, 'CAIXA'] = ""
                            df_base.at[pos, 'PALLETE'] = ""
                            df_base.at[pos, 'PLT'] = ""
                            
                            mascara_dest = (
                                (df_base["RUA"].str.strip().str.upper() == n_rua) &
                                (df_base["BOX"].str.strip().str.upper() == n_box) &
                                (df_base["ALTURA"].str.strip().str.upper() == n_alt)
                            )
                            idx_dest = df_base.index[mascara_dest].tolist()
                            
                            if idx_dest:
                                pos_dest = idx_dest[0]
                                df_base.at[pos_dest, 'CURVA'] = "C"
                                df_base.at[pos_dest, 'CODINTERNO'] = cod_int
                                df_base.at[pos_dest, 'CODFAB'] = cod_fabricante
                                df_base.at[pos_dest, 'DESCRICAO'] = desc
                                df_base.at[pos_dest, 'PALLETE'] = n_plt
                                df_base.at[pos_dest, 'PLT'] = n_plt_num
                            else:
                                novo_reg = {
                                    "CURVA": "C", "GARANTIA": "", "CODINTERNO": cod_int, "CODFAB": cod_fabricante,
                                    "DESCRICAO": desc, "CAIXA": "", "RUA": n_rua, "BOX": n_box,
                                    "ALTURA": n_alt, "PALLETE": n_plt, "PLT": n_plt_num
                                }
                                df_base = pd.concat([df_base, pd.DataFrame([novo_reg])], ignore_index=True)

                            salvar_dados(df_base)
                            registrar_historico("MOVER_PRODUTO", cod_int, desc, origem_str, destino_str, st.session_state.usuario_logado)
                            
                            st.success(f"✅ Produto {cod_fabricante} movido com sucesso para {destino_str}!")
                            st.rerun()
            else:
                st.error(f"❌ Código do fabricante '{cod_fabricante}' não localizado no estoque!")

# ============================================================
# TELA 3: CADASTRAR / OCUPAR
# ============================================================
elif opcao_menu == "➕ Cadastrar / Ocupar":
    st.header("➕ Cadastrar Produto em Endereço")
    
    if validar_admin():
        with st.form("form_cadastro_wms"):
            st.subheader("📋 Dados do Produto")
            c1, c2 = st.columns(2)
            cod_int = c1.text_input("Cód. Interno *").strip().upper()
            cod_fab = c2.text_input("Cód. Fabricante *").strip().upper()
            desc = st.text_input("Descrição Completa *").strip().upper()
            
            c3, c4 = st.columns(2)
            garantia = c3.text_input("Garantia").strip().upper()
            caixa = c4.text_input("Caixa").strip().upper()
            
            st.subheader("📍 Endereçamento Físico")
            e1, e2, e3 = st.columns(3)
            rua = e1.text_input("Rua *").strip().upper()
            box = e2.text_input("Box *").strip().upper()
            alt = e3.text_input("Altura *").strip().upper()
            
            p1, p2 = st.columns(2)
            pallet = p1.text_input("Pallet").strip().upper()
            plt_num = p2.text_input("PLT").strip().upper()
            
            btn_salvar = st.form_submit_button("💾 GRAVAR NO ESTOQUE", type="primary")
            
            if btn_salvar:
                if not (cod_int and cod_fab and desc and rua and box and alt):
                    st.error("⚠️ Preencha Cód. Interno, Cód. Fabricante, Descrição, Rua, Box e Altura!")
                else:
                    mascara = (
                        (df_base["RUA"].str.strip().str.upper() == rua) &
                        (df_base["BOX"].str.strip().str.upper() == box) &
                        (df_base["ALTURA"].str.strip().str.upper() == alt)
                    )
                    idx = df_base.index[mascara].tolist()
                    
                    novo_registro = {
                        "CURVA": "C", "GARANTIA": garantia, "CODINTERNO": cod_int, "CODFAB": cod_fab,
                        "DESCRICAO": desc, "CAIXA": caixa, "RUA": rua, "BOX": box,
                        "ALTURA": alt, "PALLETE": pallet, "PLT": plt_num
                    }
                    
                    if idx:
                        for key, val in novo_registro.items():
                            df_base.at[idx[0], key] = val
                    else:
                        df_base = pd.concat([df_base, pd.DataFrame([novo_registro])], ignore_index=True)
                        
                    salvar_dados(df_base)
                    registrar_historico("CADASTRO_PRODUTO", cod_int, desc, "NOVO CADASTRO", f"Rua:{rua} Box:{box} Alt:{alt}", st.session_state.usuario_logado)
                    st.success(f"✅ Produto '{desc}' cadastrado com sucesso na Rua {rua} / Box {box} / Altura {alt}!")

# ============================================================
# TELA 4: LIMPAR ENDEREÇO
# ============================================================
elif opcao_menu == "🧹 Limpar Endereço":
    st.header("🧹 Limpeza e Liberação de Endereço")
    
    if validar_admin():
        with st.form("form_limpeza"):
            c1, c2, c3 = st.columns(3)
            rua = c1.text_input("Rua *").strip().upper()
            box = c2.text_input("Box *").strip().upper()
            alt = c3.text_input("Altura *").strip().upper()
            
            btn_limpar = st.form_submit_button("🧹 CONFIRMAR LIMPEZA", type="primary")
            
            if btn_limpar:
                if not (rua and box and alt):
                    st.error("⚠️ Informe Rua, Box e Altura para localizar e liberar o endereço!")
                else:
                    mascara = (
                        (df_base["RUA"].str.strip().str.upper() == rua) &
                        (df_base["BOX"].str.strip().str.upper() == box) &
                        (df_base["ALTURA"].str.strip().str.upper() == alt)
                    )
                    idx = df_base.index[mascara].tolist()
                    
                    if idx:
                        cod_antigo = df_base.at[idx[0], 'CODINTERNO']
                        desc_antiga = df_base.at[idx[0], 'DESCRICAO']
                        origem = f"Rua:{rua} Box:{box} Alt:{alt}"
                        
                        df_base.at[idx[0], 'CURVA'] = ""
                        df_base.at[idx[0], 'GARANTIA'] = ""
                        df_base.at[idx[0], 'CODINTERNO'] = TEXTO_DISPONIVEL
                        df_base.at[idx[0], 'CODFAB'] = ""
                        df_base.at[idx[0], 'DESCRICAO'] = TEXTO_DISPONIVEL
                        df_base.at[idx[0], 'CAIXA'] = ""
                        df_base.at[idx[0], 'PALLETE'] = ""
                        df_base.at[idx[0], 'PLT'] = ""
                        
                        salvar_dados(df_base)
                        registrar_historico("LIMPAR_ENDERECO", cod_antigo, desc_antiga, origem, TEXTO_DISPONIVEL, st.session_state.usuario_logado)
                        st.success(f"✅ Endereço Rua {rua} / Box {box} / Altura {alt} liberado e marcado como DISPONÍVEL!")
                    else:
                        st.error("❌ Endereço informado não localizado na base!")

# ============================================================
# TELA 5: BACKUP E HISTÓRICO
# ============================================================
elif opcao_menu == "💾 Backup e Histórico":
    st.header("💾 Gestão do Banco de Dados e Histórico Auditoria")
    
    if validar_admin():
        st.subheader("📋 Histórico Completo de Movimentações e Validações")
        try:
            df_hist = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Historico", dtype=str).fillna("")
            st.dataframe(df_hist.sort_index(ascending=False), use_container_width=True, hide_index=True)
        except Exception:
            st.warning("Nenhum histórico registrado até o momento.")
            
        st.markdown("---")
        st.subheader("⚙️ Backup da Base de Dados")
        
        with open(ARQUIVO_EXCEL, "rb") as file:
            st.download_button(
                label="📥 Baixar Backup Atualizado (.xlsx)",
                data=file,
                file_name=f"Backup_WMS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )