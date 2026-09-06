import os
import time
import shutil
import winsound
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog, filedialog

# Tenta carregar bibliotecas nativas de impressão do Windows
try:
    import win32api
    import win32print
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# CONFIGURAÇÕES E CONSTANTES DO SISTEMA
ARQUIVO_EXCEL = "Estoque.xlsx"
USUARIO_PADRAO = "admin"
SENHA_PADRAO = "123"
SENHA_ADMINISTRADOR = "123"
TEXTO_DISPONIVEL = "DISPONÍVEL"

# ============================================================
# TELA DE LOGIN DO SISTEMA (WMS STYLE)
# ============================================================
class TelaLogin:
    def __init__(self, root):
        self.root = root
        self.root.title("WMS Logística - Login")
        self.root.geometry("420x280")
        self.root.configure(bg="#1E293B")
        self.root.resizable(False, False)

        self.centralizar_janela(420, 280)

        # Cabeçalho WMS
        header = tk.Frame(self.root, bg="#0F172A", height=60)
        header.pack(fill="x")
        tk.Label(header, text="🏭 WMS ESTOQUE INTELIGENTE", font=("Segoe UI", 12, "bold"), bg="#0F172A", fg="#38BDF8").pack(pady=15)

        # Formulário
        frame_form = tk.Frame(self.root, bg="#1E293B")
        frame_form.pack(pady=20)

        tk.Label(frame_form, text="Usuário:", font=("Segoe UI", 9, "bold"), bg="#1E293B", fg="#94A3B8").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_usuario = tk.Entry(frame_form, font=("Segoe UI", 11), width=22, relief="flat")
        self.ent_usuario.grid(row=0, column=1, padx=10, pady=5)
        self.ent_usuario.insert(0, "admin")

        tk.Label(frame_form, text="Senha:", font=("Segoe UI", 9, "bold"), bg="#1E293B", fg="#94A3B8").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_senha = tk.Entry(frame_form, font=("Segoe UI", 11), width=22, show="*", relief="flat")
        self.ent_senha.grid(row=1, column=1, padx=10, pady=5)
        self.ent_senha.bind("<Return>", lambda e: self.verificar_login())

        btn_entrar = tk.Button(self.root, text="🚀 ACESSAR SISTEMA WMS", bg="#0284C7", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", width=25, cursor="hand2", command=self.verificar_login)
        btn_entrar.pack(pady=10)

    def centralizar_janela(self, largura, altura):
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura // 2)
        pos_y = (altura_tela // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def verificar_login(self):
        usuario = self.ent_usuario.get().strip()
        senha = self.ent_senha.get().strip()

        if usuario == USUARIO_PADRAO and senha == SENHA_PADRAO:
            self.root.destroy()
            janela_principal = tk.Tk()

            try:
                janela_principal.state("zoomed")
            except Exception:
                largura_tela = janela_principal.winfo_screenwidth()
                altura_tela = janela_principal.winfo_screenheight()
                janela_principal.geometry(f"{largura_tela}x{altura_tela}+0+0")

            SistemaEstoque(janela_principal)
            janela_principal.mainloop()
        else:
            messagebox.showerror("Acesso Negado", "Credenciais WMS inválidas!")
            self.ent_senha.delete(0, tk.END)
            self.ent_senha.focus()


# ============================================================
# SISTEMA PRINCIPAL WMS DE ESTOQUE
# ============================================================
class SistemaEstoque:
    def __init__(self, root):
        self.root = root
        self.root.title("WMS - Sistema de Gestão e Conferência de Estoque")
        self.root.configure(bg="#F1F5F9")

        self.modo_congelado = False
        self.processando_validacao = False

        self.inicializar_excel()
        self.configurar_estilos()

        self.verificar_atualizacao_sexta()

        # --- PAINEL SUPERIOR WMS ---
        frame_topo = tk.LabelFrame(self.root, text=" Painel de Operações WMS ", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#0F172A", padx=15, pady=10, relief="solid", bd=1)
        frame_topo.pack(fill="x", padx=15, pady=10)

        tk.Label(frame_topo, text="Busca (F1):", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#334155").grid(row=0, column=0, sticky="w", padx=5)
        self.ent_busca = tk.Entry(frame_topo, width=22, font=("Segoe UI", 11), relief="solid", bd=1)
        self.ent_busca.grid(row=0, column=1, padx=5, pady=5)
        self.ent_busca.bind("<Return>", lambda e: self.pesquisar_produto())

        tk.Button(frame_topo, text="🔍 BUSCAR", bg="#0284C7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.pesquisar_produto).grid(row=0, column=2, padx=3)

        self.btn_congelar = tk.Button(frame_topo, text="❄️ CONGELAR: OFF", bg="#64748B", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.alternar_congelar)
        self.btn_congelar.grid(row=0, column=3, padx=3)

        tk.Button(frame_topo, text="🟢 DISPONÍVEIS", bg="#0D9488", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.buscar_disponiveis).grid(row=0, column=4, padx=3)

        tk.Button(frame_topo, text="📊 CALCULAR CURVA ABC", bg="#0284C7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.calcular_curva_abc_por_pesquisa).grid(row=0, column=5, padx=3)

        tk.Button(frame_topo, text="🖨️ IMPRIMIR", bg="#D97706", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.imprimir_resultados).grid(row=0, column=6, padx=3)

        tk.Label(frame_topo, text="Leitor (F2):", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#334155").grid(row=1, column=0, sticky="w", padx=5, pady=8)
        self.ent_validacao = tk.Entry(frame_topo, width=22, font=("Segoe UI", 11), relief="solid", bd=1)
        self.ent_validacao.grid(row=1, column=1, padx=5, pady=8)
        self.ent_validacao.bind("<Return>", lambda e: self.validar_produto())

        self.lbl_status = tk.Label(frame_topo, text="⏸️ AGUARDANDO", font=("Segoe UI", 10, "bold"), bg="#E2E8F0", fg="#334155", width=18, height=1, relief="solid", bd=1)
        self.lbl_status.grid(row=1, column=2, columnspan=2, padx=5, sticky="w")

        # --- PAINEL DE COMANDOS WMS (ADMIN) ---
        frame_admin = tk.LabelFrame(self.root, text=" Operações de Armazém WMS (Requer Senha Admin) ", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#B91C1C", padx=15, pady=8, relief="solid", bd=1)
        frame_admin.pack(fill="x", padx=15, pady=0)

        tk.Button(frame_admin, text="📦 MOVER", bg="#4F46E5", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.mover_produto).grid(row=0, column=0, padx=4, pady=2)
        tk.Button(frame_admin, text="🧹 LIMPAR", bg="#DC2626", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.limpar_endereco).grid(row=0, column=1, padx=4, pady=2)
        tk.Button(frame_admin, text="➕ CADASTRAR", bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.cadastrar_produto).grid(row=0, column=2, padx=4, pady=2)
        tk.Button(frame_admin, text="🔄 ATUALIZAR EM MASSA", bg="#7C3AED", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.atualizar_em_massa).grid(row=0, column=3, padx=4, pady=2)
        tk.Button(frame_admin, text="💾 BACKUP", bg="#0284C7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, cursor="hand2", command=self.gerar_backup_base).grid(row=0, column=4, padx=4, pady=2)

        # --- TABELA DE RESULTADOS WMS ---
        frame_tabela = tk.Frame(self.root, bg="#F1F5F9")
        frame_tabela.pack(fill="both", expand=True, padx=15, pady=10)

        self.colunas = ("CURVA", "GARANTIA", "COD. INTERNO", "COD. FABRICANTE", "DESCRIÇÃO DO PRODUTO", "CAIXA", "RUA", "BOX", "ALTURA", "PALLETE", "PLT")
        self.tabela = ttk.Treeview(frame_tabela, columns=self.colunas, show="headings", selectmode="browse")

        larguras = [65, 85, 120, 130, 280, 70, 60, 60, 65, 80, 60]
        for idx, col in enumerate(self.colunas):
            self.tabela.heading(col, text=col)
            self.tabela.column(col, width=larguras[idx], anchor="center")
        self.tabela.column("DESCRIÇÃO DO PRODUTO", anchor="w")

        scroll_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tabela.yview)
        scroll_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tabela.pack(side="top", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")

        # --- BARRA DE STATUS (RODAPÉ) ---
        frame_rodape = tk.Frame(self.root, bg="#E2E8F0", height=25)
        frame_rodape.pack(fill="x", side="bottom")

        self.lbl_rodape = tk.Label(frame_rodape, text="WMS Pronto | Base Conectada: Estoque.xlsx", font=("Segoe UI", 8, "bold"), bg="#E2E8F0", fg="#475569")
        self.lbl_rodape.pack(side="left", padx=10)

        self.root.bind("<F1>", lambda e: self.ent_busca.focus())
        self.root.bind("<F2>", lambda e: self.ent_validacao.focus())

    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#1E293B", foreground="white", borderwidth=1, relief="flat")
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26, background="white", fieldbackground="white")
        style.map("Treeview", background=[("selected", "#0284C7")], foreground=[("selected", "white")])

    def inicializar_excel(self):
        cols_base = ["CURVA", "GARANTIA", "CODINTERNO", "CODFAB", "DESCRICAO", "CAIXA", "RUA", "BOX", "ALTURA", "PALLETE", "PLT"]
        cols_hist = ["DATA_HORA", "ACAO", "CODINTERNO", "DESCRICAO", "ORIGEM", "DESTINO"]
        cols_pesquisa = ["DATA_HORA", "TERMO_BUSCADO", "CODFAB_ENCONTRADO"]

        if not os.path.exists(ARQUIVO_EXCEL):
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
                pd.DataFrame(columns=cols_base).to_excel(writer, sheet_name="Base_Dados", index=False)
                pd.DataFrame(columns=cols_hist).to_excel(writer, sheet_name="Historico", index=False)
                pd.DataFrame(columns=cols_pesquisa).to_excel(writer, sheet_name="Controle_Pesquisas", index=False)
            return

        # O arquivo já existe (pode ter sido criado por uma versão mais antiga
        # do sistema, sem todas as abas). Antes, o código só verificava a
        # coluna CURVA em Base_Dados e presumia que Historico e
        # Controle_Pesquisas já existiam — se não existissem, qualquer
        # tentativa de ler/gravar nelas quebrava com "Worksheet ... not found".
        # Agora conferimos e criamos cada aba que estiver faltando.
        try:
            planilhas_existentes = pd.ExcelFile(ARQUIVO_EXCEL, engine="openpyxl").sheet_names
        except Exception:
            planilhas_existentes = []

        abas_padrao = {
            "Base_Dados": pd.DataFrame(columns=cols_base),
            "Historico": pd.DataFrame(columns=cols_hist),
            "Controle_Pesquisas": pd.DataFrame(columns=cols_pesquisa),
        }
        abas_faltando = {
            nome: df_vazio
            for nome, df_vazio in abas_padrao.items()
            if nome not in planilhas_existentes
        }

        if abas_faltando:
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a") as writer:
                for nome_aba, df_vazio in abas_faltando.items():
                    df_vazio.to_excel(writer, sheet_name=nome_aba, index=False)

        # Garante que a coluna CURVA exista na Base_Dados
        try:
            df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str)
            if "CURVA" not in df_base.columns:
                df_base.insert(0, "CURVA", "")
                with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_base.to_excel(writer, sheet_name="Base_Dados", index=False)
        except Exception:
            pass

    def verificar_atualizacao_sexta(self):
        if datetime.now().weekday() == 4:
            hoje_str = datetime.now().strftime("%Y-%m-%d")
            arquivo_controle = "ultima_atualizacao_abc.txt"
            rodou_hoje = False

            if os.path.exists(arquivo_controle):
                with open(arquivo_controle, "r") as f:
                    if f.read().strip() == hoje_str:
                        rodou_hoje = True

            if not rodou_hoje:
                self.processar_calculo_abc_silencioso()
                with open(arquivo_controle, "w") as f:
                    f.write(hoje_str)

    # --------------------------------------------------------
    # CURVA ABC (LÓGICA CENTRAL — CORRIGIDA)
    # --------------------------------------------------------
    def _calcular_curva_abc(self):
        """
        Calcula a curva ABC dos produtos com base no volume de pesquisas
        registrado em Controle_Pesquisas e grava o resultado na coluna
        CURVA da planilha Base_Dados.

        Retorna uma tupla (sucesso: bool, mensagem: str).
        """
        df_p = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Controle_Pesquisas", dtype=str).fillna("")

        # Remove pesquisas sem código de fabricante válido. Sem isso, uma
        # chave vazia ("") no value_counts() poderia ser casada, no merge,
        # com QUALQUER produto que também estivesse com CODFAB em branco,
        # inflando a contagem desses produtos de forma incorreta.
        df_p = df_p[df_p["CODFAB_ENCONTRADO"].str.strip() != ""]

        if df_p.empty:
            return False, "Ainda não há registros de pesquisas suficientes para calcular a curva!"

        contagem = df_p["CODFAB_ENCONTRADO"].value_counts().reset_index()
        contagem.columns = ["CODFAB", "TOTAL_PESQUISAS"]

        df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")

        # Separa os endereços vazios/disponíveis dos produtos reais ANTES de
        # calcular os cortes. Antes, os endereços "DISPONÍVEL" ficavam
        # misturados na mesma lista ordenada usada para contar as posições,
        # o que deslocava o corte A/B/C dos produtos reais.
        mascara_vazio = (
            (df_base["CODINTERNO"].str.strip() == TEXTO_DISPONIVEL)
            | (df_base["CODINTERNO"].str.strip() == "")
        )
        df_itens = df_base[~mascara_vazio].copy()
        df_vazios = df_base[mascara_vazio].copy()

        if df_itens.empty:
            return False, "Não há produtos cadastrados para calcular a curva ABC!"

        df_itens = df_itens.merge(contagem, on="CODFAB", how="left")
        df_itens["TOTAL_PESQUISAS"] = df_itens["TOTAL_PESQUISAS"].fillna(0)
        df_itens = df_itens.sort_values(by="TOTAL_PESQUISAS", ascending=False).reset_index(drop=True)

        total_itens = len(df_itens)

        # Antes, com poucos produtos cadastrados (ex.: 4 itens), int(4*0.20)
        # e int(4*0.30) arredondavam para 0, então NENHUM produto virava "A"
        # ou "B" e tudo caía em "C" — mesmo os mais pesquisados. Usamos
        # round() e garantimos ao menos 1 item por faixa (quando possível).
        corte_a = max(1, round(total_itens * 0.20))
        corte_b = corte_a + max(1, round(total_itens * 0.30))
        corte_b = min(corte_b, total_itens)  # nunca deixa a faixa B "vazar" além do total

        def classificar(posicao):
            if posicao < corte_a:
                return "A"
            elif posicao < corte_b:
                return "B"
            return "C"

        df_itens["CURVA"] = [classificar(i) for i in range(total_itens)]
        df_itens = df_itens.drop(columns=["TOTAL_PESQUISAS"])

        df_vazios = df_vazios.copy()
        df_vazios["CURVA"] = ""

        df_final = pd.concat([df_itens, df_vazios], ignore_index=True)
        df_final = df_final[df_base.columns]  # preserva a ordem original das colunas

        with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_final.to_excel(writer, sheet_name="Base_Dados", index=False)

        qtd_a = min(corte_a, total_itens)
        qtd_b = corte_b - corte_a
        qtd_c = total_itens - corte_b
        mensagem = (
            "Curva ABC calculada com sucesso com base no volume de pesquisas!\n\n"
            f"A: {qtd_a} produto(s) | B: {qtd_b} produto(s) | C: {qtd_c} produto(s)"
        )
        return True, mensagem

    def calcular_curva_abc_por_pesquisa(self):
        if not self.autenticar_admin(): return
        try:
            sucesso, mensagem = self._calcular_curva_abc()

            if not sucesso:
                messagebox.showwarning("Aviso WMS", mensagem)
                return

            messagebox.showinfo("Sucesso WMS", mensagem)
            self.lbl_rodape.config(text="WMS Curva ABC Atualizada | Base Conectada")

            # Atualiza a tabela em tela imediatamente para refletir a nova
            # curva. Antes, a planilha era recalculada mas a tela ficava
            # com os dados antigos até uma nova busca manual, parecendo
            # que o cálculo "não tinha feito nada".
            termo_atual = self.ent_busca.get().strip()
            if termo_atual:
                self.pesquisar_produto()
            elif self.modo_congelado:
                pass  # em modo congelado, preserva o que já está acumulado na tela
            else:
                for item in self.tabela.get_children():
                    self.tabela.delete(item)

        except Exception as e:
            messagebox.showerror("Erro WMS", f"Não foi possível calcular a curva:\n{e}")

    def processar_calculo_abc_silencioso(self):
        try:
            self._calcular_curva_abc()
        except Exception:
            pass

    def autenticar_admin(self):
        senha = simpledialog.askstring("WMS Segurança", "Digite a senha de administrador:", show="*", parent=self.root)
        if senha == SENHA_ADMINISTRADOR:
            return True
        elif senha is not None:
            messagebox.showerror("Acesso Negado", "Senha incorreta!")
        return False

    def registrar_historico(self, acao, cod, desc, orig, dest):
        try:
            df_hist = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Historico", dtype=str).fillna("")
            novo_log = {
                "DATA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ACAO": acao, "CODINTERNO": cod, "DESCRICAO": desc, "ORIGEM": orig, "DESTINO": dest
            }
            df_hist = pd.concat([df_hist, pd.DataFrame([novo_log])], ignore_index=True)
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_hist.to_excel(writer, sheet_name="Historico", index=False)
        except Exception as e:
            print("Erro ao registrar histórico:", e)

    def registrar_pesquisa(self, termo, cod_fab_encontrado):
        try:
            df_p = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Controle_Pesquisas", dtype=str).fillna("")
            novo_p = {
                "DATA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "TERMO_BUSCADO": termo,
                "CODFAB_ENCONTRADO": cod_fab_encontrado
            }
            df_p = pd.concat([df_p, pd.DataFrame([novo_p])], ignore_index=True)
            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_p.to_excel(writer, sheet_name="Controle_Pesquisas", index=False)
        except Exception as e:
            print("Erro ao registrar pesquisa:", e)

    def exibir_alerta_visual(self, mensagem, eh_valido):
        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)

        largura, altura = 420, 140
        pos_x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (largura // 2)
        pos_y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (altura // 2)
        top.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

        cor_fundo = "#16A34A" if eh_valido else "#DC2626"
        top.configure(bg=cor_fundo)

        lbl = tk.Label(top, text=f"{'✅' if eh_valido else '❌'} {mensagem}", font=("Segoe UI", 24, "bold"), fg="white", bg=cor_fundo)
        lbl.pack(expand=True)

        self.root.update()
        time.sleep(1.2)
        top.destroy()

    def validar_produto(self):
        if self.processando_validacao: return
        self.processando_validacao = True

        cod_bipado = self.ent_validacao.get().strip().upper()
        if not cod_bipado:
            self.processando_validacao = False
            return

        itens_na_tela = self.tabela.get_children()
        encontrado_na_tela = False

        for item in itens_na_tela:
            valores = self.tabela.item(item, "values")
            cod_interno_tela = str(valores[2]).strip().upper()
            cod_fab_tela = str(valores[3]).strip().upper()

            if cod_bipado == cod_interno_tela or cod_bipado == cod_fab_tela:
                encontrado_na_tela = True
                break

        if encontrado_na_tela:
            self.lbl_status.config(text="✅ VALIDADO", bg="#DCFCE7", fg="#16A34A")
            try:
                winsound.Beep(1200, 150)
            except Exception:
                pass
            self.exibir_alerta_visual("VALIDADO", True)
        else:
            self.lbl_status.config(text="❌ NÃO VALIDADO", bg="#FEE2E2", fg="#DC2626")
            try:
                winsound.Beep(400, 250)
                winsound.Beep(300, 200)
            except Exception:
                pass
            self.exibir_alerta_visual("NÃO VALIDADO", False)

        self.ent_validacao.delete(0, tk.END)
        self.ent_validacao.focus()
        self.processando_validacao = False

    def pesquisar_produto(self):
        termo = self.ent_busca.get().strip().upper()
        if not termo:
            messagebox.showwarning("Aviso", "Digite o código ou descrição para pesquisar!")
            return

        df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")
        resultados = df_base[
            (df_base["CODFAB"].str.upper().str.contains(termo)) |
            (df_base["CODINTERNO"].str.upper().str.contains(termo)) |
            (df_base["DESCRICAO"].str.upper().str.contains(termo)) |
            (df_base["GARANTIA"].str.upper().str.contains(termo))
        ]
        resultados = resultados[~resultados["CODINTERNO"].str.upper().str.contains("DISPON")]

        for _, row in resultados.iterrows():
            self.registrar_pesquisa(termo, row["CODFAB"])

        if not self.modo_congelado:
            for item in self.tabela.get_children(): self.tabela.delete(item)

        self.tabela.tag_configure('CURVA_A', background='#DCFCE7')
        self.tabela.tag_configure('CURVA_B', background='#FEF9C3')
        self.tabela.tag_configure('CURVA_C', background='#FFFFFF')

        for idx, row in resultados.iterrows():
            if len(self.tabela.get_children()) >= 50:
                break
            curva = str(row.get("CURVA", "C")).strip().upper()
            tag = f"CURVA_{curva}" if curva in ["A", "B", "C"] else "CURVA_C"

            self.tabela.insert("", "end", values=tuple(row), tags=(tag,))

        total = len(self.tabela.get_children())
        self.lbl_rodape.config(text=f"Registros carregados na tela: {total} | WMS Ativo")

        if resultados.empty and not self.modo_congelado:
            messagebox.showinfo("Sem Resultados", f"Nenhum produto encontrado para: {termo}")
        elif self.modo_congelado:
            self.ent_busca.delete(0, tk.END)

    def alternar_congelar(self):
        self.modo_congelado = not self.modo_congelado
        if self.modo_congelado:
            self.btn_congelar.config(text="❄️ CONGELAR: ON", bg="#16A34A")
            messagebox.showinfo("WMS Acumulativo", "MODO CONGELADO ATIVADO!")
        else:
            self.btn_congelar.config(text="❄️ CONGELAR: OFF", bg="#64748B")
            for item in self.tabela.get_children(): self.tabela.delete(item)
            self.ent_busca.delete(0, tk.END)
            self.lbl_rodape.config(text="WMS Limpo | Base Conectada")

    def buscar_disponiveis(self):
        df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")
        resultados = df_base[(df_base["CODINTERNO"].str.upper().str.contains("DISPON")) | (df_base["CODINTERNO"] == "")]

        for item in self.tabela.get_children(): self.tabela.delete(item)

        self.tabela.tag_configure('CURVA_A', background='#DCFCE7')
        self.tabela.tag_configure('CURVA_B', background='#FEF9C3')
        self.tabela.tag_configure('CURVA_C', background='#FFFFFF')

        for _, row in resultados.iterrows():
            if len(self.tabela.get_children()) >= 50: break
            self.tabela.insert("", "end", values=tuple(row), tags=("CURVA_C",))

        self.lbl_rodape.config(text=f"Endereços disponíveis listados: {len(self.tabela.get_children())}")

    def imprimir_resultados(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Não há dados na tela para imprimir!")
            return

        arquivo_temp = "relatorio_wms.txt"
        with open(arquivo_temp, "w", encoding="utf-8") as f:
            f.write("RELATÓRIO WMS - CONFERÊNCIA DE ESTOQUE\n")
            f.write("Data/Hora: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "\n")
            f.write("=" * 65 + "\n\n")
            for child in itens:
                v = self.tabela.item(child, "values")
                f.write(f"CURVA: {v[0]} | GARANTIA: {v[1]} | COD.INT: {v[2]} | COD.FAB: {v[3]}\n")
                f.write(f"DESC: {v[4]}\n")
                f.write(f"RUA: {v[6]} | BOX: {v[7]} | ALTURA: {v[8]} | CX: {v[5]} | PLT: {v[10]}\n")
                f.write("-" * 65 + "\n")

        try:
            if HAS_WIN32:
                win32api.ShellExecute(0, "printto", arquivo_temp, f'"{win32print.GetDefaultPrinter()}"', ".", 0)
                messagebox.showinfo("WMS Impressão", "Relatório enviado com sucesso para a impressora padrão!")
            else:
                os.startfile(arquivo_temp, "print")
        except Exception:
            try:
                os.startfile(arquivo_temp, "print")
                messagebox.showinfo("WMS Impressão", "Enviado para a impressora com sucesso!")
            except Exception as ex:
                messagebox.showerror("Erro de Impressão", f"Falha ao comunicar com a impressora corporativa:\n{ex}")

    def atualizar_em_massa(self):
        if not self.autenticar_admin(): return

        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione a planilha para Importação em Massa WMS",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Arquivos CSV", "*.csv")]
        )
        if not caminho_arquivo: return

        try:
            if caminho_arquivo.endswith('.csv'):
                df_novo = pd.read_csv(caminho_arquivo, dtype=str).fillna("")
            else:
                df_novo = pd.read_excel(caminho_arquivo, dtype=str).fillna("")

            cols_obrigatorias = ["CURVA", "GARANTIA", "CODINTERNO", "CODFAB", "DESCRICAO", "CAIXA", "RUA", "BOX", "ALTURA", "PALLETE", "PLT"]

            for col in ["CODINTERNO", "DESCRICAO"]:
                if col not in df_novo.columns:
                    messagebox.showerror("Erro WMS", f"O arquivo selecionado não contém a coluna obrigatória: {col}")
                    return

            for col in cols_obrigatorias:
                if col not in df_novo.columns:
                    df_novo[col] = ""

            df_novo = df_novo[cols_obrigatorias]

            with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_novo.to_excel(writer, sheet_name="Base_Dados", index=False)

            self.registrar_historico("IMPORTACAO_MASSA", "MULTIPLOS", f"{len(df_novo)} registros", "ARQUIVO EXTERNO", "BASE WMS")
            messagebox.showinfo("Sucesso WMS", f"Importação em massa concluída com sucesso!\n\nTotal de {len(df_novo)} linhas atualizadas na base.")
            self.lbl_rodape.config(text=f"WMS Atualizado | Total de registros: {len(df_novo)}")

        except Exception as e:
            messagebox.showerror("Erro de Importação", f"Não foi possível processar o arquivo:\n{e}")

    def mover_produto(self):
        if not self.autenticar_admin(): return

        top = tk.Toplevel(self.root)
        top.title("WMS - Mover Produto")
        top.geometry("380x350")
        top.configure(bg="#F1F5F9")
        top.resizable(False, False)
        top.grab_set()

        tk.Label(top, text="📦 MOVIMENTAÇÃO DE ARMAZÉM", font=("Segoe UI", 10, "bold"), bg="#F1F5F9", fg="#0F172A").pack(pady=12)

        form = tk.Frame(top, bg="#F1F5F9")
        form.pack(padx=20, pady=5)

        tk.Label(form, text="Cód. Fabricante:", font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=0, column=0, sticky="w", pady=3)
        ent_fab = tk.Entry(form, font=("Segoe UI", 10), width=22)
        ent_fab.grid(row=0, column=1, padx=5, pady=3)
        ent_fab.focus()

        tk.Label(form, text="Nova Rua:", font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=1, column=0, sticky="w", pady=3)
        ent_rua = tk.Entry(form, font=("Segoe UI", 10), width=22)
        ent_rua.grid(row=1, column=1, padx=5, pady=3)

        tk.Label(form, text="Novo Box:", font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=2, column=0, sticky="w", pady=3)
        ent_box = tk.Entry(form, font=("Segoe UI", 10), width=22)
        ent_box.grid(row=2, column=1, padx=5, pady=3)

        tk.Label(form, text="Nova Altura:", font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=3, column=0, sticky="w", pady=3)
        ent_alt = tk.Entry(form, font=("Segoe UI", 10), width=22)
        ent_alt.grid(row=3, column=1, padx=5, pady=3)

        tk.Label(form, text="Novo Pallet:", font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=4, column=0, sticky="w", pady=3)
        ent_plt_desc = tk.Entry(form, font=("Segoe UI", 10), width=22)
        ent_plt_desc.grid(row=4, column=1, padx=5, pady=3)

        tk.Label(form, text="Novo PLT:", font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=5, column=0, sticky="w", pady=3)
        ent_plt_num = tk.Entry(form, font=("Segoe UI", 10), width=22)
        ent_plt_num.grid(row=5, column=1, padx=5, pady=3)

        def salvar_movimento():
            cod_fabricante = ent_fab.get().strip()
            nova_rua = ent_rua.get().strip()
            novo_box = ent_box.get().strip()
            nova_alt = ent_alt.get().strip()
            novo_pallet = ent_plt_desc.get().strip()
            novo_plt = ent_plt_num.get().strip()

            if not cod_fabricante:
                messagebox.showwarning("Aviso", "Informe o Código do Fabricante!", parent=top)
                return

            df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")
            idx = df_base.index[df_base["CODFAB"].str.strip().str.upper() == cod_fabricante.strip().upper()].tolist()

            if idx:
                origem = f"Rua:{df_base.at[idx[0], 'RUA']} Box:{df_base.at[idx[0], 'BOX']} Alt:{df_base.at[idx[0], 'ALTURA']}"
                destino = f"Rua:{nova_rua} Box:{novo_box} Alt:{nova_alt} Plt:{novo_plt}"

                df_base.at[idx[0], 'RUA'] = nova_rua
                df_base.at[idx[0], 'BOX'] = novo_box
                df_base.at[idx[0], 'ALTURA'] = nova_alt
                df_base.at[idx[0], 'PALLETE'] = novo_pallet
                df_base.at[idx[0], 'PLT'] = novo_plt

                with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_base.to_excel(writer, sheet_name="Base_Dados", index=False)

                cod_interno = df_base.at[idx[0], 'CODINTERNO']
                self.registrar_historico("MOVER", cod_interno, df_base.at[idx[0], 'DESCRICAO'], origem, destino)
                messagebox.showinfo("Sucesso WMS", "Produto movimentado com sucesso!", parent=top)
                top.destroy()
            else:
                messagebox.showerror("Erro WMS", "Código do Fabricante não encontrado na base!", parent=top)

        tk.Button(top, text="💾 SALVAR MOVIMENTAÇÃO", bg="#4F46E5", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", width=24, cursor="hand2", command=salvar_movimento).pack(pady=10)

    def cadastrar_produto(self):
        if not self.autenticar_admin(): return

        top = tk.Toplevel(self.root)
        top.title("WMS - Cadastrar Produto")
        top.geometry("380x420")
        top.configure(bg="#F1F5F9")
        top.resizable(False, False)
        top.grab_set()

        tk.Label(top, text="➕ CADASTRO DE NOVO PRODUTO", font=("Segoe UI", 10, "bold"), bg="#F1F5F9", fg="#0F172A").pack(pady=12)

        form = tk.Frame(top, bg="#F1F5F9")
        form.pack(padx=20, pady=5)

        campos = ["Cód. Interno:", "Cód. Fabricante:", "Descrição:", "Garantia:", "Caixa:", "Rua:", "Box:", "Altura:", "Pallet:", "PLT:"]
        entradas = {}

        for i, campo in enumerate(campos):
            tk.Label(form, text=campo, font=("Segoe UI", 9, "bold"), bg="#F1F5F9", fg="#334155").grid(row=i, column=0, sticky="w", pady=2)
            ent = tk.Entry(form, font=("Segoe UI", 10), width=22)
            ent.grid(row=i, column=1, padx=5, pady=2)
            entradas[campo] = ent

        def salvar_cadastro():
            novo_dado = {
                "CURVA": "C",
                "GARANTIA": entradas["Garantia:"].get().strip(),
                "CODINTERNO": entradas["Cód. Interno:"].get().strip(),
                "CODFAB": entradas["Cód. Fabricante:"].get().strip(),
                "DESCRICAO": entradas["Descrição:"].get().strip(),
                "CAIXA": entradas["Caixa:"].get().strip(),
                "RUA": entradas["Rua:"].get().strip(),
                "BOX": entradas["Box:"].get().strip(),
                "ALTURA": entradas["Altura:"].get().strip(),
                "PALLETE": entradas["Pallet:"].get().strip(),
                "PLT": entradas["PLT:"].get().strip()
            }

            if not novo_dado["CODINTERNO"] or not novo_dado["DESCRICAO"]:
                messagebox.showwarning("Aviso", "Cód. Interno e Descrição são obrigatórios!", parent=top)
                return

            try:
                df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")
                df_base = pd.concat([df_base, pd.DataFrame([novo_dado])], ignore_index=True)

                with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_base.to_excel(writer, sheet_name="Base_Dados", index=False)

                self.registrar_historico("CADASTRAR", novo_dado["CODINTERNO"], novo_dado["DESCRICAO"], "EXTERNAS", "ESTOQUE")
                messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso!", parent=top)
                top.destroy()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao cadastrar produto:\n{e}", parent=top)

        tk.Button(top, text="💾 SALVAR CADASTRO", bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", width=24, cursor="hand2", command=salvar_cadastro).pack(pady=10)

    def limpar_endereco(self):
        if not self.autenticar_admin(): return
        cod_fab = simpledialog.askstring("WMS Limpar", "Digite o Código do Fabricante que deseja zerar/remover do endereço:", parent=self.root)
        if not cod_fab: return

        try:
            df_base = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base_Dados", dtype=str).fillna("")
            idx = df_base.index[df_base["CODFAB"].str.strip().str.upper() == cod_fab.strip().upper()].tolist()

            if idx:
                i = idx[0]
                desc = df_base.at[i, "DESCRICAO"]
                cod_int = df_base.at[i, "CODINTERNO"]

                df_base.at[i, "CODINTERNO"] = TEXTO_DISPONIVEL
                df_base.at[i, "GARANTIA"] = ""
                df_base.at[i, "DESCRICAO"] = ""
                df_base.at[i, "CURVA"] = ""

                with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_base.to_excel(writer, sheet_name="Base_Dados", index=False)

                self.registrar_historico("LIMPAR_ENDERECO", cod_int, desc, "ENDEREÇO", "DISPONÍVEL")
                messagebox.showinfo("Sucesso WMS", f"Endereço do produto {cod_fab} liberado como DISPONÍVEL!")
            else:
                messagebox.showerror("Erro", "Código do Fabricante não encontrado na base!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao limpar endereço:\n{e}")

    def gerar_backup_base(self):
        if not self.autenticar_admin(): return
        try:
            data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_backup = f"Backup_Estoque_{data_hora}.xlsx"
            shutil.copy(ARQUIVO_EXCEL, nome_backup)
            messagebox.showinfo("Backup WMS", f"Backup gerado com sucesso!\nArquivo salvo como: {nome_backup}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível gerar o backup:\n{e}")


# ============================================================
# EXECUÇÃO DO APLICATIVO
# ============================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = TelaLogin(root)
    root.mainloop()