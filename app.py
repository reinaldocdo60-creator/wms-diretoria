import os
import json
import io
from datetime import datetime
import pandas as pd
import streamlit as str_lit

# =========================================================
# CONFIGURAÇÕES INICIAIS DA PÁGINA
# =========================================================
str_lit.set_page_config(
    page_title="WMS Litle - Sistema de Gestão de Armazém",
    page_icon="📦",
    layout="wide"
)

# =========================================================
# ARQUIVOS DE PERSISTÊNCIA
# =========================================================
ARQUIVO_EXCEL = "Base_Estoque.xlsx"
ARQUIVO_USUARIOS = "usuarios_wms.json"

# =========================================================
# FUNÇÕES DE USUÁRIOS E AUTENTICAÇÃO
# =========================================================
def carregar_usuarios():
    if os.path.exists(ARQUIVO_USUARIOS):
        try:
            with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "admin": {"senha": "admin123", "perfil": "ADMIN"},
        "operador": {"senha": "123", "perfil": "OPERADOR"}
    }

def salvar_usuarios(usuarios):
    with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=4)

if "usuarios_db" not in str_lit.session_state:
    str_lit.session_state["usuarios_db"] = carregar_usuarios()

# Garantia de inicialização segura das variáveis de sessão
if "autenticado" not in str_lit.session_state:
    str_lit.session_state["autenticado"] = False
if "usuario_atual" not in str_lit.session_state:
    str_lit.session_state["usuario_atual"] = ""
if "perfil_atual" not in str_lit.session_state:
    str_lit.session_state["perfil_atual"] = ""

# =========================================================
# FUNÇÕES DE DADOS (WMS)
# =========================================================
def carregar_dados():
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
            df.columns = [c.strip().upper() for c in df.columns]
            return df
        except Exception as e:
            str_lit.error(f"Erro ao carregar o arquivo Excel: {e}")
    
    dados_exemplo = {
        "GARANTIA": ["GARANTIA A", "GARANTIA B", "GARANTIA A"],
        "CODIGO": ["PROD001", "PROD002", "PROD003"],
        "DESCRICAO": ["Amortecedor Dianteiro", "Pastilha de Freio", "Filtro de Óleo"],
        "RUA": ["R01", "R02", "R03"],
        "BOX": ["B05", "B12", "B01"],
        "ALTURA": ["A1", "A2", "A3"],
        "QUANTIDADE": ["10", "25", "50"]
    }
    df_ex = pd.DataFrame(dados_exemplo)
    df_ex.to_excel(ARQUIVO_EXCEL, index=False)
    return df_ex

def salvar_dados(df):
    df.to_excel(ARQUIVO_EXCEL, index=False)

if "df_estoque" not in str_lit.session_state:
    str_lit.session_state["df_estoque"] = carregar_dados()

if "linhas_congeladas" not in str_lit.session_state:
    str_lit.session_state["linhas_congeladas"] = pd.DataFrame()

# =========================================================
# ESTILOS CSS E LAYOUT DE IMPRESSÃO PROFISSIONAL
# =========================================================
CSS_IMPRESSAO = """
<style>
@media print {
    header, [data-testid="stSidebar"], .stButton, .stTextInput, .stSelectbox, .stAlert, .stFileUploader {
        display: none !important;
    }
    body {
        background-color: white !important;
        color: black !important;
        font-family: Arial, sans-serif;
    }
    .report-container {
        width: 100%;
        margin: 0;
        padding: 20px;
    }
    .report-header {
        border-bottom: 2px solid #333;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    .report-footer {
        margin-top: 40px;
        border-top: 1px solid #ccc;
        padding-top: 10px;
        font-size: 10px;
        color: #555;
        text-align: right;
    }
}
</style>
"""
str_lit.markdown(CSS_IMPRESSAO, unsafe_allow_html=True)

# =========================================================
# TELA DE LOGIN
# =========================================================
if not str_lit.session_state["autenticado"]:
    str_lit.markdown("<h1 style='text-align: center;'>📦 WMS Litle - Login</h1>", unsafe_allow_html=True)
    
    col1, col2, col3 = str_lit.columns([1, 1.2, 1])
    with col2:
        with str_lit.form("form_login"):
            usuario_input = str_lit.text_input("Usuário").strip().lower()
            senha_input = str_lit.text_input("Senha", type="password").strip()
            btn_login = str_lit.form_submit_button("Entrar no Sistema", use_container_width=True)
            
            if btn_login:
                usuarios_db = str_lit.session_state["usuarios_db"]
                if usuario_input in usuarios_db and usuarios_db[usuario_input]["senha"] == senha_input:
                    str_lit.session_state["autenticado"] = True
                    str_lit.session_state["usuario_atual"] = usuario_input
                    str_lit.session_state["perfil_atual"] = usuarios_db[usuario_input]["perfil"]
                    str_lit.success("Login realizado com sucesso!")
                    str_lit.rerun()
                else:
                    str_lit.error("Usuário ou senha incorretos.")
    str_lit.stop()

# =========================================================
# BARRA LATERAL (MENU E PERFIL)
# =========================================================
with str_lit.sidebar:
    str_lit.image("https://cdn-icons-png.flaticon.com/512/2821/2821870.png", width=70)
    nome_exibicao = str_lit.session_state.get('usuario_atual', 'Usuário').capitalize()
    perfil_exibicao = str_lit.session_state.get('perfil_atual', 'OPERADOR')
    str_lit.markdown(f"### Olá, **{nome_exibicao}**")
    str_lit.markdown(f"Perfil: `👤 {perfil_exibicao}`")
    str_lit.markdown("---")
    
    opcao_menu = str_lit.radio(
        "Navegação Principal",
        [
            "🔍 Pesquisa e Validação",
            "🚚 Mover Produto",
            "➕ Cadastrar / Ocupar",
            "📥 Importar / Atualizar Base",
            "📊 Sugestão de Inventário",
            "💾 Backup e Histórico",
            "📖 Manual de Instruções",
            "👥 Gerenciar Usuários"
        ]
    )
    
    str_lit.markdown("---")
    if str_lit.button("🚪 Sair (Logout)", use_container_width=True):
        str_lit.session_state["autenticado"] = False
        str_lit.session_state["usuario_atual"] = ""
        str_lit.session_state["perfil_atual"] = ""
        str_lit.rerun()

def validar_admin():
    if str_lit.session_state.get("perfil_atual", "") != "ADMIN":
        str_lit.warning("⚠️ Acesso restrito! Esta função exige perfil de Administrador.")
        return False
    return True

df = str_lit.session_state["df_estoque"]

# =========================================================
# TELA 1: PESQUISA E VALIDAÇÃO
# =========================================================
if opcao_menu == "🔍 Pesquisa e Validação":
    
    if str_lit.session_state.get("modo_impressao_congelados", False):
        usuario_rel = str_lit.session_state.get('usuario_atual', 'admin').capitalize()
        str_lit.markdown(f"""
        <div class="report-container">
            <div class="report-header">
                <h2>📦 WMS Litle - Relatório de Itens Congelados</h2>
                <p><b>Usuário Emitente:</b> {usuario_rel}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        str_lit.info("💡 Pressione **Ctrl + P** no seu teclado para enviar diretamente à impressora ou salvar em PDF.")
        
        if str_lit.button("⬅️ Voltar para a Pesquisa", use_container_width=True):
            str_lit.session_state["modo_impressao_congelados"] = False
            str_lit.rerun()
            
        str_lit.markdown("---")
        if not str_lit.session_state["linhas_congeladas"].empty:
            str_lit.dataframe(str_lit.session_state["linhas_congeladas"], use_container_width=True)
        else:
            str_lit.warning("Não há linhas congeladas para exibir.")
            
        str_lit.markdown(f"""
        <div class="report-footer">
            Relatório gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} | WMS Litle - Sistema de Gestão de Armazém
        </div>
        """, unsafe_allow_html=True)
        str_lit.stop()

    str_lit.header("🔍 Pesquisa e Validação de Estoque")
    
    if "GARANTIA" in df.columns:
        garantias_disponiveis = ["TODAS"] + sorted(list(df["GARANTIA"].dropna().unique()))
    else:
        garantias_disponiveis = ["TODAS"]

    c1, c2, c3 = str_lit.columns([2, 3, 2])
    with c1:
        filtro_garantia = str_lit.selectbox("Filtrar por Garantia", garantias_disponiveis)
    with c2:
        termo_busca = str_lit.text_input("Digite o Código, Descrição ou Endereço").strip()
    with c3:
        termo_validacao = str_lit.text_input("Validação por Bipe/Código").strip()

    df_filtrado = df.copy()
    if filtro_garantia != "TODAS":
        df_filtrado = df_filtrado[df_filtrado["GARANTIA"] == filtro_garantia]

    if termo_busca:
        mascara = df_filtrado.astype(str).apply(lambda x: x.str.contains(termo_busca, case=False, na=False)).any(axis=1)
        df_filtrado = df_filtrado[mascara]

    str_lit.subheader("📋 Resultados da Pesquisa")
    if not df_filtrado.empty:
        str_lit.dataframe(df_filtrado, use_container_width=True)
        
        if str_lit.session_state.get("perfil_atual", "") == "ADMIN":
            str_lit.markdown("---")
            str_lit.markdown("### 🗑️ Gerenciamento Rápido (Exclusão por Linha)")
            
            opcoes_linhas = []
            for idx, row in df_filtrado.iterrows():
                garantia_val = row.get("GARANTIA", "N/D")
                codigo_val = row.get("CODIGO", "N/D")
                desc_val = row.get("DESCRICAO", "N/D")
                rua_val = row.get("RUA", "N/D")
                box_val = row.get("BOX", "N/D")
                texto_exibicao = f"[{garantia_val}] Código: {codigo_val} | Descrição: {desc_val} | Endereço: Rua {rua_val}, Box {box_val} (Índice Base: {idx})"
                opcoes_linhas.append((idx, texto_exibicao))
            
            selecao_excluir = str_lit.selectbox(
                "Selecione um registro dos resultados acima para excluir permanentemente:",
                options=[None] + [item[0] for item in opcoes_linhas],
                format_func=lambda x: next((item[1] for item in opcoes_linhas if item[0] == x), "Selecione um item...") if x is not None else "Selecione um item..."
            )
            
            if selecao_excluir is not None:
                if str_lit.button("❌ Excluir Esta Linha Selecionada", type="primary"):
                    df = df.drop(index=selecao_excluir).reset_index(drop=True)
                    salvar_dados(df)
                    str_lit.session_state["df_estoque"] = df
                    str_lit.success("✅ Registro excluído com sucesso do banco de dados!")
                    str_lit.rerun()

        col_cong1, col_cong2 = str_lit.columns(2)
        with col_cong1:
            if str_lit.button("📌 Congelar Seleção Atual", use_container_width=True):
                novos_congelados = pd.concat([str_lit.session_state["linhas_congeladas"], df_filtrado]).drop_duplicates()
                str_lit.session_state["linhas_congeladas"] = novos_congelados
                str_lit.success("Linhas congeladas com sucesso!")
        with col_cong2:
            if str_lit.button("🗑️ Limpar Congelados", use_container_width=True):
                str_lit.session_state["linhas_congeladas"] = pd.DataFrame(columns=df.columns)
                str_lit.success("Lista congelada limpa!")
    else:
        str_lit.info("Nenhum registro encontrado com os filtros informados.")

    if not str_lit.session_state["linhas_congeladas"].empty:
        str_lit.markdown("---")
        col_t1, col_t2 = str_lit.columns([3, 1])
        with col_t1:
            str_lit.subheader("📌 Itens Congelados / Salvos Temporariamente")
        with col_t2:
            if str_lit.button("🖨️ Imprimir Congelados", use_container_width=True):
                str_lit.session_state["modo_impressao_congelados"] = True
                str_lit.rerun()
                
        str_lit.dataframe(str_lit.session_state["linhas_congeladas"], use_container_width=True)

    if termo_validacao:
        encontrado = df[df.astype(str).apply(lambda x: x.str.contains(termo_validacao, case=False, na=False)).any(axis=1)]
        if not encontrado.empty:
            str_lit.success(f"✅ Peça Validada com Sucesso! Encontrada na base.")
        else:
            str_lit.error(f"❌ Atenção! Código '{termo_validacao}' não localizado na base de estoque.")

# =========================================================
# TELA 2: MOVER PRODUTO
# =========================================================
elif opcao_menu == "🚚 Mover Produto":
    str_lit.header("🚚 Movimentação de Endereço de Produto")
    
    with str_lit.form("form_mover"):
        codigo_mover = str_lit.text_input("Código do Produto a Mover").strip().upper()
        
        c_r, c_b, c_a = str_lit.columns(3)
        with c_r:
            nova_rua = str_lit.text_input("Nova Rua").strip().upper()
        with c_b:
            novo_box = str_lit.text_input("Novo Box").strip().upper()
        with c_a:
            nova_altura = str_lit.text_input("Nova Altura").strip().upper()
            
        btn_executar_movimento = str_lit.form_submit_button("Atualizar Endereço", use_container_width=True)
        
        if btn_executar_movimento:
            if not codigo_mover or not nova_rua or not novo_box or not nova_altura:
                str_lit.warning("Preencha todos os campos obrigatórios para movimentação.")
            else:
                if "CODIGO" in df.columns:
                    mask = df["CODIGO"] == codigo_mover
                    if mask.any():
                        df.loc[mask, "RUA"] = nova_rua
                        df.loc[mask, "BOX"] = novo_box
                        df.loc[mask, "ALTURA"] = nova_altura
                        salvar_dados(df)
                        str_lit.session_state["df_estoque"] = df
                        str_lit.success(f"✅ Produto {codigo_mover} movimentado com sucesso para Rua: {nova_rua}, Box: {novo_box}, Altura: {nova_altura}!")
                    else:
                        str_lit.error(f"❌ Código {codigo_mover} não encontrado na base.")

# =========================================================
# TELA 3: CADASTRAR / OCUPAR
# =========================================================
elif opcao_menu == "➕ Cadastrar / Ocupar":
    str_lit.header("➕ Cadastrar Novo Item ou Ocupar Endereço")
    if validar_admin():
        with str_lit.form("form_cadastrar"):
            c_g, c_c = str_lit.columns(2)
            with c_g:
                cad_garantia = str_lit.text_input("Garantia").strip().upper()
            with c_c:
                cad_codigo = str_lit.text_input("Código do Produto").strip().upper()
                
            cad_descricao = str_lit.text_input("Descrição do Produto").strip()
            
            cc1, cc2, cc3, cc4 = str_lit.columns(4)
            with cc1:
                cad_rua = str_lit.text_input("Rua").strip().upper()
            with cc2:
                cad_box = str_lit.text_input("Box").strip().upper()
            with cc3:
                cad_altura = str_lit.text_input("Altura").strip().upper()
            with cc4:
                cad_qtd = str_lit.text_input("Quantidade").strip()
                
            btn_salvar_cadastro = str_lit.form_submit_button("Salvar Novo Cadastro", use_container_width=True)
            
            if btn_salvar_cadastro:
                if not cad_codigo or not cad_rua:
                    str_lit.warning("Informe pelo menos o Código e a Rua.")
                else:
                    nova_linha = {
                        "GARANTIA": cad_garantia,
                        "CODIGO": cad_codigo,
                        "DESCRICAO": cad_descricao,
                        "RUA": cad_rua,
                        "BOX": cad_box,
                        "ALTURA": cad_altura,
                        "QUANTIDADE": cad_qtd
                    }
                    df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
                    salvar_dados(df)
                    str_lit.session_state["df_estoque"] = df
                    str_lit.success("✅ Item cadastrado com sucesso na base!")

# =========================================================
# TELA 4: IMPORTAR / ATUALIZAR BASE
# =========================================================
elif opcao_menu == "📥 Importar / Atualizar Base":
    str_lit.header("📥 Importação de Nova Base de Estoque")
    if validar_admin():
        str_lit.write("Faça o upload de uma nova planilha Excel (`.xlsx`) para atualizar a base de dados central.")
        
        output = io.BytesIO()
        df_modelo = pd.DataFrame(columns=["GARANTIA", "CODIGO", "DESCRICAO", "RUA", "BOX", "ALTURA", "QUANTIDADE"])
        df_modelo.loc[0] = ["GARANTIA A", "EXEMPLO01", "Peça Exemplo", "R01", "B01", "A1", "10"]
        df_modelo.to_excel(output, index=False)
        output.seek(0)
        
        str_lit.download_button(
            label="📥 Baixar Planilha Modelo Oficial (Gabarito)",
            data=output,
            file_name="Modelo_Importacao_WMS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        str_lit.markdown("---")
        arquivo_submetido = str_lit.file_uploader("Escolha o arquivo Excel preenchido", type=["xlsx"])
        if arquivo_submetido is not None:
            try:
                df_novo = pd.read_excel(arquivo_submetido, dtype=str)
                df_novo.columns = [c.strip().upper() for c in df_novo.columns]
                
                str_lit.write("Pré-visualização dos novos dados:")
                str_lit.dataframe(df_novo.head(), use_container_width=True)
                
                if str_lit.button("💾 Substituir Base de Dados Oficial", type="primary", use_container_width=True):
                    df_novo.to_excel(ARQUIVO_EXCEL, index=False)
                    str_lit.session_state["df_estoque"] = df_novo
                    str_lit.success("✅ Base de dados substituída e atualizada com sucesso!")
            except Exception as e:
                str_lit.error(f"Erro ao processar o arquivo enviado: {e}")

# =========================================================
# TELA 5: SUGESTÃO DE INVENTÁRIO
# =========================================================
elif opcao_menu == "📊 Sugestão de Inventário":
    str_lit.header("📊 Sugestão de Inventário Cíclico")
    str_lit.write("Abaixo está uma amostragem orientada para auditoria preventiva do armazém.")
    
    if not df.empty:
        qtd_amostra = min(10, len(df))
        df_sugestao_10 = df.sample(n=qtd_amostra)
        str_lit.table(df_sugestao_10)
    else:
        str_lit.warning("Nenhum dado disponível na base para gerar sugestão de inventário.")

# =========================================================
# TELA 6: BACKUP E HISTÓRICO
# =========================================================
elif opcao_menu == "💾 Backup e Histórico":
    str_lit.header("💾 Backup e Histórico da Base de Dados")
    str_lit.write("Faça o download do arquivo de backup atual da base de estoque do WMS em formato Excel (.xlsx).")
    
    if os.path.exists(ARQUIVO_EXCEL):
        with open(ARQUIVO_EXCEL, "rb") as file:
            str_lit.download_button(
                label="📥 Baixar Backup Atual (Base_Estoque.xlsx)",
                data=file,
                file_name=f"Backup_WMS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        str_lit.warning("Nenhum arquivo de base de estoque encontrado para backup.")

# =========================================================
# TELA 7: MANUAL DE INSTRUÇÕES
# =========================================================
elif opcao_menu == "📖 Manual de Instruções":
    
    if str_lit.session_state.get("modo_impressao_manual", False):
        usuario_rel = str_lit.session_state.get('usuario_atual', 'admin').capitalize()
        str_lit.markdown(f"""
        <div class="report-container">
            <div class="report-header">
                <h2>📖 WMS Litle - Manual de Instruções e Operação</h2>
                <p><b>Usuário Emitente:</b> {usuario_rel}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        str_lit.info("💡 Pressione **Ctrl + P** no seu teclado para enviar diretamente à impressora ou salvar em PDF.")
        
        if str_lit.button("⬅️ Voltar para o Manual Normal", use_container_width=True):
            str_lit.session_state["modo_impressao_manual"] = False
            str_lit.rerun()
            
        str_lit.markdown("---")
        str_lit.markdown("""
        ### WMS LITLE - MANUAL RÁPIDO DE OPERAÇÃO
        1. **Pesquisa e Validação:** Utilize os filtros por Garantia e termos de busca para localizar rapidamente itens e endereços. Use o campo de validação para conferir o código bipeado da peça separada. Os administradores contam com exclusão rápida por linha integrada.
        2. **Mover Produto:** Altere a localização de itens informando o código e o novo endereço de rua, box e altura.
        3. **Cadastrar / Ocupar:** Adicione novos itens à base preenchendo todos os campos obrigatórios (Restrito a Admin).
        4. **Importar / Atualizar Base:** Baixe a planilha modelo oficial com os cabeçalhos corretos (`GARANTIA`, `CODIGO`, `DESCRICAO`, `RUA`, `BOX`, `ALTURA`, `QUANTIDADE`) e faça a substituição em massa via arquivo Excel (Restrito a Admin).
        5. **Backup e Gerenciamento:** Realize backups de segurança e gerencie usuários e perfis com facilidade.
        """)
        
        str_lit.markdown(f"""
        <div class="report-footer">
            Manual impresso em {datetime.now().strftime('%d/%m/%Y às %H:%M')} | WMS Litle - Sistema de Gestão de Armazém
        </div>
        """, unsafe_allow_html=True)
        str_lit.stop()

    str_lit.header("📖 Manual de Instruções do WMS")
    
    col_man1, col_man2 = str_lit.columns([3, 1])
    with col_man1:
        str_lit.info("💡 Consulte abaixo as instruções rápidas de utilização de cada módulo do sistema.")
    with col_man2:
        if str_lit.button("🖨️ Visualizar / Imprimir", use_container_width=True):
            str_lit.session_state["modo_impressao_manual"] = True
            str_lit.rerun()

    str_lit.markdown("""
    ### 📌 Guia de Utilização - WMS LITLE

    * **🔍 Pesquisa e Validação (Geral):**
      * Selecione a **Garantia** desejada para filtrar o escopo inicial.
      * Digite no campo de **Pesquisa** o código, descrição ou endereço.
      * Utilize o campo de **Validação** para bipar/digitar o código da peça separada e confirmar se ela pertence ao grupo consultado.
      * Administradores possuem um menu em cascata logo abaixo dos resultados para **excluir individualmente** qualquer linha indesejada direto no banco.
      * É possível **Congelar** linhas da pesquisa e imprimir o relatório formatado.

    * **🚚 Mover Produto:**
      * Informe o código interno ou do fabricante do produto.
      * Preencha os novos dados de **Rua**, **Box** e **Altura** para atualizar o mapeamento em tempo real.

    * **➕ Cadastrar / Ocupar (ADMIN):**
      * Permite adicionar novos registros e endereçar itens na base de dados do WMS.

    * **📥 Importar / Atualizar Base em Massa (ADMIN):**
      * **Novo:** Disponibiliza o botão para **Baixar a Planilha Modelo Oficial** com os títulos exatos (`GARANTIA`, `CODIGO`, `DESCRICAO`, `RUA`, `BOX`, `ALTURA`, `QUANTIDADE`) exigidos pelo sistema.
      * Permite o upload do arquivo preenchido para substituir e atualizar instantaneamente toda a base central.

    * **📊 Sugestão de Inventário Cíclico:**
      * Gera uma amostragem inicial orientada para auditoria preventiva.

    * **💾 Backup e Histórico:**
      * Permite baixar a cópia de segurança atual da base de dados em formato `.xlsx`.

    * **👥 Gerenciar Usuários (ADMIN):**
      * Permite cadastrar novos operadores e administradores ou remover acessos existentes.
    """)

# =========================================================
# TELA 8: GERENCIAR USUÁRIOS (APENAS ADMIN)
# =========================================================
elif opcao_menu == "👥 Gerenciar Usuários":
    str_lit.header("👥 Gerenciamento de Usuários e Senhas")
    if validar_admin():
        usuarios_db = str_lit.session_state["usuarios_db"]
        
        str_lit.subheader("📋 Usuários Cadastrados Atualmente")
        df_usuarios = pd.DataFrame([
            {"USUARIO": u.capitalize(), "PERFIL": info["perfil"], "SENHA": info["senha"]}
            for u, info in usuarios_db.items()
        ])
        str_lit.table(df_usuarios)
        
        str_lit.markdown("---")
        str_lit.subheader("➕ Cadastrar ou Atualizar Usuário")
        
        with str_lit.form("form_novo_usuario"):
            novo_user_nome = str_lit.text_input("Nome de Usuário (Login)").strip().lower()
            novo_user_senha = str_lit.text_input("Senha").strip()
            novo_user_perfil = str_lit.selectbox("Perfil de Acesso", ["OPERADOR", "ADMIN"])
            
            btn_salvar_usuario = str_lit.form_submit_button("Salvar Usuário", use_container_width=True)
            
            if btn_salvar_usuario:
                if not novo_user_nome or not novo_user_senha:
                    str_lit.warning("Preencha o nome de usuário e a senha.")
                else:
                    usuarios_db[novo_user_nome] = {
                        "senha": novo_user_senha,
                        "perfil": novo_user_perfil
                    }
                    salvar_usuarios(usuarios_db)
                    str_lit.success(f"✅ Usuário '{novo_user_nome.capitalize()}' salvo/atualizado com sucesso!")
                    str_lit.rerun()

        str_lit.markdown("---")
        str_lit.subheader("🗑️ Remover Usuário")
        usuarios_removiveis = [u for u in usuarios_db.keys() if u != "admin"]
        if usuarios_removiveis:
            user_para_remover = str_lit.selectbox("Selecione o usuário para remover", usuarios_removiveis)
            if str_lit.button("❌ Excluir Usuário", type="primary"):
                if user_para_remover in usuarios_db:
                    del usuarios_db[user_para_remover]
                    salvar_usuarios(usuarios_db)
                    str_lit.success(f"✅ Usuário '{user_para_remover.capitalize()}' removido com sucesso!")
                    str_lit.rerun()
        else:
            str_lit.info("Não há outros usuários removíveis além do administrador principal.")
