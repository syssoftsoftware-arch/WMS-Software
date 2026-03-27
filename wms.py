import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
import pyodbc, os, pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- ESTILIZAÇÃO ---
VERSION = "v12.5.1"
COR_P, COR_S, COR_A, COR_E = "#0078D7", "#28a745", "#FFCC00", "#8E44AD"


class DBManager:
    def __init__(self):
        # Altere para o caminho da sua rede se necessário
        self.db_path = os.path.abspath("wms_db.accdb")
        self.conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                         f'DBQ={self.db_path};Exclusive=No;Threads=4;')

    def executar(self, query, params=(), fetch=False):
        conn = None
        try:
            conn = pyodbc.connect(self.conn_str, timeout=10)
            cursor = conn.cursor()
            cursor.execute(query, params)
            if fetch:
                res = cursor.fetchall()
                return [list(row) for row in res] if res else []
            conn.commit()
            return True
        except Exception as e:
            print(f"Erro DB: {e}");
            return [] if fetch else False
        finally:
            if conn: conn.close()


# --- LOGIN ---
class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = DBManager()
        self.title("WMS Enterprise - Login")
        self.geometry("400x500")
        ctk.CTkLabel(self, text="WMS LOGÍSTICA 4.0", font=("Arial", 24, "bold"), text_color=COR_P).pack(pady=50)
        self.u = ctk.CTkEntry(self, placeholder_text="Usuário", width=280, height=40);
        self.u.pack(pady=10)
        self.s = ctk.CTkEntry(self, placeholder_text="Senha", show="*", width=280, height=40);
        self.s.pack(pady=10)
        ctk.CTkButton(self, text="ENTRAR", fg_color=COR_P, width=280, height=45, command=self.logar).pack(pady=30)

    def logar(self):
        res = self.db.executar("SELECT [Login], [Perfil] FROM [Usuarios] WHERE [Login]=? AND [Senha]=?",
                               (self.u.get(), self.s.get()), fetch=True)
        if res:
            self.destroy()
            WMSApp(res[0][0], res[0][1]).mainloop()
        else:
            messagebox.showerror("Erro", "Login Inválido")


# --- APP PRINCIPAL ---
class WMSApp(ctk.CTk):
    def __init__(self, user, perfil):
        super().__init__()
        self.user, self.perfil, self.db = user, perfil, DBManager()
        self.dados_nota_cega = {}
        self.title(f"WMS v{VERSION} - {self.user.upper()}");
        self.geometry("1300x850")

        self.tabs = ctk.CTkTabview(self, segmented_button_selected_color=COR_P)
        self.tabs.pack(padx=20, pady=10, fill="both", expand=True)

        self.tab_rec = self.tabs.add("📥 RECEBIMENTO")
        self.tab_inv = self.tabs.add("📦 ESTOQUE")
        self.tab_exp = self.tabs.add("🚚 EXPEDIÇÃO")
        self.tab_dash = self.tabs.add("📊 DASHBOARD")
        self.tab_hist = self.tabs.add("📜 AUDITORIA")

        if self.perfil == "ADM":
            self.tab_adm = self.tabs.add("⚙️ ADM");
            self.setup_adm()

        self.setup_recebimento();
        self.setup_estoque();
        self.setup_expedicao()
        self.setup_dashboard();
        self.setup_historico()

    # --- LÓGICA DE INTELIGÊNCIA (SLOTTING) ---
    def obter_curva_abc(self, sku):
        """Define Curva A, B ou C baseada no giro histórico"""
        res = self.db.executar("SELECT COUNT(*) FROM [Historico] WHERE [SKU]=? AND [Tipo_Mov]='SAIDA'", (sku,),
                               fetch=True)
        giro = res[0][0] if res else 0
        if giro > 50: return "A"
        if giro > 15: return "B"
        return "C"

    def sugerir_endereco(self, sku):
        """Busca o melhor local vazio para a curva do produto"""
        curva = self.obter_curva_abc(sku)
        sugestao = self.db.executar(
            "SELECT TOP 1 [Endereco] FROM [Mapa_Armazem] WHERE [Curva_Alvo]=? AND [Status]='Vazio' ORDER BY [Distancia_Docas] ASC",
            (curva,), fetch=True)
        return sugestao[0][0] if sugestao else "PULMAO_GERAL"

    # --- INTERFACE: RECEBIMENTO ---
    def setup_recebimento(self):
        f_top = ctk.CTkFrame(self.tab_rec);
        f_top.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(f_top, text="📂 LER XML NF-e", command=self.importar_xml).pack(side="left", padx=10)
        self.lbl_nf = ctk.CTkLabel(f_top, text="Aguardando XML...", text_color="gray");
        self.lbl_nf.pack(side="left", padx=10)

        f_body = ctk.CTkFrame(self.tab_rec);
        f_body.pack(fill="both", expand=True)
        self.f_in = ctk.CTkFrame(f_body, width=300);
        self.f_in.pack(side="left", fill="y", padx=10, pady=10)

        self.c_sku = ctk.CTkEntry(self.f_in, placeholder_text="SKU");
        self.c_sku.pack(pady=5, padx=10, fill="x")
        self.c_qtd = ctk.CTkEntry(self.f_in, placeholder_text="Quantidade");
        self.c_qtd.pack(pady=5, padx=10, fill="x")
        self.c_lote = ctk.CTkEntry(self.f_in, placeholder_text="Lote");
        self.c_lote.pack(pady=5, padx=10, fill="x")
        self.c_val = ctk.CTkEntry(self.f_in, placeholder_text="Validade");
        self.c_val.pack(pady=5, padx=10, fill="x")
        self.c_st = ctk.CTkComboBox(self.f_in, values=["Disponível", "Quarentena", "Avariado"]);
        self.c_st.pack(pady=5, padx=10, fill="x");
        self.c_st.set("Disponível")

        ctk.CTkButton(self.f_in, text="✔️ VALIDAR ENTRADA", fg_color=COR_S, command=self.processar_entrada).pack(
            pady=20, padx=10, fill="x")
        self.tree_rec = self.criar_tree(f_body, ("SKU", "Descrição", "Esperado", "Lido", "Divergência"))

    def importar_xml(self):
        path = filedialog.askopenfilename(filetypes=[("XML", "*.xml")])
        if not path: return
        try:
            tree = ET.parse(path);
            root = tree.getroot()
            self.dados_nota_cega = {}
            for tag in root.iter():
                if tag.tag.endswith('nNF'): n_nf = tag.text
            for det in root.findall('.//{*}det'):
                sku = det.find('.//{*}cProd').text
                nome = det.find('.//{*}xProd').text
                qtd = int(float(det.find('.//{*}qCom').text))
                self.dados_nota_cega[sku] = {"desc": nome, "qtd_esp": qtd, "lido": 0}
            self.lbl_nf.configure(text=f"NF: {n_nf} | Itens Carregados", text_color="cyan");
            self.refresh_rec()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro XML: {e}")

    def processar_entrada(self):
        sku, q = self.c_sku.get(), self.c_qtd.get()
        if sku not in self.dados_nota_cega: return messagebox.showerror("Erro", "SKU inválido")
        try:
            qtd = int(q);
            self.dados_nota_cega[sku]["lido"] += qtd
            # INTELIGÊNCIA: Sugere endereço antes de gravar
            end_sugerido = self.sugerir_endereco(sku)

            self.db.executar(
                "INSERT INTO [Estoque] ([SKU],[Produto],[Qtd],[Endereco],[Lote],[Validade],[Status]) VALUES (?,?,?,?,?,?,?)",
                (sku, self.dados_nota_cega[sku]["desc"], qtd, end_sugerido, self.c_lote.get(), self.c_val.get(),
                 self.c_st.get()))

            self.db.executar("UPDATE [Mapa_Armazem] SET [Status]='Ocupado' WHERE [Endereco]=?", (end_sugerido,))
            self.registrar_log(sku, f"ENTRADA_{self.c_st.get()}", qtd, end_sugerido)

            messagebox.showinfo("Endereçamento",
                                f"SUGESTÃO: Guardar no endereço {end_sugerido}\n(Curva {self.obter_curva_abc(sku)})")
            self.refresh_rec();
            self.refresh_inv();
            self.limpar_rec()
        except:
            messagebox.showerror("Erro", "Verifique as quantidades")

    # --- ESTOQUE E EXPEDIÇÃO ---
    def setup_estoque(self):
        self.tree_inv = self.criar_tree(self.tab_inv,
                                        ("SKU", "Produto", "Qtd", "Endereço", "Status", "Lote", "Validade"))
        ctk.CTkButton(self.tab_inv, text="↻ ATUALIZAR", command=self.refresh_inv).pack(pady=5);
        self.refresh_inv()

    def setup_expedicao(self):
        f = ctk.CTkFrame(self.tab_exp);
        f.pack(pady=10, fill="x", padx=10)
        self.e_ped = ctk.CTkEntry(f, placeholder_text="Pedido Saída");
        self.e_ped.pack(side="left", padx=10)
        ctk.CTkButton(f, text="BAIXAR ITEM", fg_color=COR_E, command=self.baixar_item).pack(side="right", padx=10)
        self.tree_exp = self.criar_tree(self.tab_exp, ("SKU", "Produto", "Qtd", "Local", "Status"));
        self.refresh_exp()

    def baixar_item(self):
        sel = self.tree_exp.selection()
        if not sel or not self.e_ped.get(): return
        it = self.tree_exp.item(sel)['values']
        if it[4] != "Disponível": return messagebox.showerror("Erro", "Bloqueio de Qualidade")

        # Trava de Segurança Rede
        if self.db.executar("DELETE FROM (SELECT TOP 1 * FROM [Estoque] WHERE [SKU]=? AND [Status]='Disponível')",
                            (str(it[0]),)):
            self.db.executar("UPDATE [Mapa_Armazem] SET [Status]='Vazio' WHERE [Endereco]=?", (it[3],))
            self.registrar_log(it[0], "SAIDA", it[2], f"PED:{self.e_ped.get()}")
            self.refresh_inv();
            self.refresh_exp()

    # --- DASHBOARD E AUDITORIA ---
    def setup_dashboard(self):
        self.f_dash = ctk.CTkFrame(self.tab_dash);
        self.f_dash.pack(fill="both", expand=True, padx=20, pady=20)
        self.render_kpi()

    def render_kpi(self):
        for w in self.f_dash.winfo_children(): w.destroy()
        dados = self.db.executar("SELECT [Status], SUM([Qtd]) FROM [Estoque] GROUP BY [Status]", fetch=True)
        if not dados: return
        fig, ax = plt.subplots(figsize=(5, 3));
        fig.patch.set_facecolor('#242424');
        ax.set_facecolor('#242424')
        ax.pie([float(d[1]) for d in dados], labels=[d[0] for d in dados], colors=[COR_S, COR_A, "#dc3545"],
               textprops={'color': "w"})
        FigureCanvasTkAgg(fig, master=self.f_dash).get_tk_widget().pack(fill="both", expand=True)

    def setup_historico(self):
        ctk.CTkButton(self.tab_hist, text="EXPORTAR EXCEL", fg_color="#1D6F42",
                      command=lambda: pd.DataFrame(self.db.executar("SELECT * FROM [Historico]", fetch=True)).to_excel(
                          "WMS_Auditoria.xlsx")).pack(pady=10)
        self.tree_hist = self.criar_tree(self.tab_hist, ("ID", "SKU", "Tipo", "Qtd", "Destino", "User", "Data"));
        self.refresh_hist()

    # --- HELPERS ---
    def setup_adm(self):
        f = ctk.CTkFrame(self.tab_adm);
        f.pack(pady=20, fill="x", padx=20)
        self.nu, self.ns = ctk.CTkEntry(f, placeholder_text="Login"), ctk.CTkEntry(f, placeholder_text="Senha")
        self.nu.pack(side="left", padx=5);
        self.ns.pack(side="left", padx=5)
        ctk.CTkButton(f, text="+ CADASTRAR", command=lambda: self.db.executar(
            "INSERT INTO [Usuarios] ([Login],[Senha],[Perfil]) VALUES (?,?,?)",
            (self.nu.get(), self.ns.get(), "OPERADOR"))).pack(side="left")

    def registrar_log(self, sku, t, q, d):
        self.db.executar(
            "INSERT INTO [Historico] ([SKU],[Tipo_Mov],[Quantidade],[Origem_Destino],[Usuario],[Data_Hora]) VALUES (?,?,?,?,?,?)",
            (sku, t, int(q), d, self.user, datetime.now().strftime("%d/%m/%Y %H:%M")))
        self.refresh_hist()

    def criar_tree(self, master, cols):
        t = ttk.Treeview(master, columns=cols, show="headings");
        t.pack(fill="both", expand=True, padx=10, pady=10)
        for c in cols: t.heading(c, text=c); t.column(c, width=120, anchor="center")
        return t

    def refresh_rec(self):
        for i in self.tree_rec.get_children(): self.tree_rec.delete(i)
        for s, info in self.dados_nota_cega.items(): self.tree_rec.insert("", "end",
                                                                          values=(s, info["desc"], "???", info["lido"],
                                                                                  info["lido"] - info["qtd_esp"]))

    def refresh_inv(self):
        for i in self.tree_inv.get_children(): self.tree_inv.delete(i)
        for r in self.db.executar("SELECT [SKU],[Produto],[Qtd],[Endereco],[Status],[Lote],[Validade] FROM [Estoque]",
                                  fetch=True): self.tree_inv.insert("", "end", values=tuple(r))

    def refresh_exp(self):
        for i in self.tree_exp.get_children(): self.tree_exp.delete(i)
        for r in self.db.executar("SELECT [SKU],[Produto],[Qtd],[Endereco],[Status] FROM [Estoque]",
                                  fetch=True): self.tree_exp.insert("", "end", values=tuple(r))

    def refresh_hist(self):
        for i in self.tree_hist.get_children(): self.tree_hist.delete(i)
        for r in self.db.executar("SELECT * FROM [Historico] ORDER BY [ID] DESC", fetch=True): self.tree_hist.insert("",
                                                                                                                     "end",
                                                                                                                     values=tuple(
                                                                                                                         r))

    def limpar_rec(self):
        for e in [self.c_sku, self.c_qtd, self.c_lote, self.c_val]: e.delete(0, 'end')
        self.c_sku.focus()


if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    if os.path.exists("wms_db.accdb"):
        LoginWindow().mainloop()
    else:
        messagebox.showerror("Erro", "Banco 'wms_db.accdb' não encontrado!")