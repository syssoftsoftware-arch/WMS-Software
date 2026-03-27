import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import os

# --- CONFIGURAÇÕES DE ACESSO ---
USUARIO_ADM = "admin"
SENHA_ADM = "1234"


# --- TELA DE LOGIN ---
class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Acesso Restrito - WMS")
        self.geometry("400x350")
        self.resizable(False, False)

        # Centralização visual
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="SISTEMA WMS PRO", font=("Arial", 24, "bold")).grid(row=0, column=0, pady=(30, 20))

        self.user_entry = ctk.CTkEntry(self, placeholder_text="Usuário", width=250)
        self.user_entry.grid(row=1, column=0, pady=10)

        self.pass_entry = ctk.CTkEntry(self, placeholder_text="Senha", show="*", width=250)
        self.pass_entry.grid(row=2, column=0, pady=10)

        self.btn_login = ctk.CTkButton(self, text="Entrar no Sistema", command=self.verificar_login, width=250,
                                       height=40)
        self.btn_login.grid(row=3, column=0, pady=30)

    def verificar_login(self):
        if self.user_entry.get() == USUARIO_ADM and self.pass_entry.get() == SENHA_ADM:
            self.destroy()  # Fecha login
            app = WMSPro()  # Abre WMS
            app.mainloop()
        else:
            messagebox.showerror("Erro de Acesso", "Usuário ou Senha incorretos!")


# --- SISTEMA WMS PRINCIPAL ---
class WMSPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"WMS Pro v6.0 - Logado como: {USUARIO_ADM}")
        self.geometry("1100x700")

        # Configurações de Arquivo e Regras
        self.arquivo_excel = "estoque_wms.xlsx"
        self.arquivo_reposicao = "reposicao_necessaria.xlsx"
        self.limite_minimo = 5

        self.carregar_dados()
        self.setup_ui()
        self.mostrar_estoque()

    def carregar_dados(self):
        colunas = ['SKU', 'Produto', 'Qtd', 'Endereco', 'Data']
        if os.path.exists(self.arquivo_excel):
            self.estoque = pd.read_excel(self.arquivo_excel)
            for col in colunas:
                if col not in self.estoque.columns: self.estoque[col] = ""
        else:
            self.estoque = pd.DataFrame(columns=colunas)

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Sidebar (Esquerda) ---
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=3, sticky="nsew")

        ctk.CTkLabel(self.sidebar, text="OPERAÇÕES", font=("Arial", 18, "bold")).pack(pady=20)

        self.entry_sku = ctk.CTkEntry(self.sidebar, placeholder_text="SKU do Item")
        self.entry_sku.pack(padx=20, pady=5, fill="x")

        self.entry_nome = ctk.CTkEntry(self.sidebar, placeholder_text="Nome do Produto")
        self.entry_nome.pack(padx=20, pady=5, fill="x")

        self.entry_qtd = ctk.CTkEntry(self.sidebar, placeholder_text="Quantidade")
        self.entry_qtd.pack(padx=20, pady=5, fill="x")

        self.entry_local = ctk.CTkEntry(self.sidebar, placeholder_text="Endereço (Ex: R-02-B)")
        self.entry_local.pack(padx=20, pady=5, fill="x")

        self.btn_entrada = ctk.CTkButton(self.sidebar, text="📥 Registrar Entrada", fg_color="#2980B9",
                                         command=self.registrar_entrada)
        self.btn_entrada.pack(padx=20, pady=20, fill="x")

        self.btn_saida = ctk.CTkButton(self.sidebar, text="📤 Registrar Saída (Baixa)", fg_color="#C0392B",
                                       command=self.registrar_saida)
        self.btn_saida.pack(padx=20, pady=5, fill="x")

        self.btn_relatorio = ctk.CTkButton(self.sidebar, text="📊 Relatório de Compras", fg_color="#27AE60",
                                           command=self.gerar_relatorio_reposicao)
        self.btn_relatorio.pack(padx=20, pady=(50, 10), fill="x")

        # --- Painel de Busca (Topo) ---
        self.frame_busca = ctk.CTkFrame(self, height=70)
        self.frame_busca.grid(row=0, column=1, padx=20, pady=10, sticky="ew")

        self.entry_busca = ctk.CTkEntry(self.frame_busca, placeholder_text="Pesquisar por SKU, Nome ou Localização...",
                                        width=500)
        self.entry_busca.pack(side="left", padx=15, pady=15, expand=True, fill="x")

        self.btn_buscar = ctk.CTkButton(self.frame_busca, text="🔍 Buscar", command=self.filtrar_busca, width=100)
        self.btn_buscar.pack(side="left", padx=5)

        self.btn_refresh = ctk.CTkButton(self.frame_busca, text="🔄", width=40, fg_color="gray",
                                         command=self.mostrar_estoque)
        self.btn_refresh.pack(side="left", padx=5)

        # --- Área de Visualização (Centro) ---
        self.textbox = ctk.CTkTextbox(self, font=("Courier New", 13))
        self.textbox.grid(row=1, column=1, padx=20, pady=10, sticky="nsew")

        # Alertas (Rodapé)
        self.lbl_alerta = ctk.CTkLabel(self, text="", font=("Arial", 12, "bold"))
        self.lbl_alerta.grid(row=2, column=1, pady=10)

    # --- LÓGICA DE NEGÓCIO ---
    def registrar_entrada(self):
        try:
            sku, nome, qtd, local = self.entry_sku.get(), self.entry_nome.get(), int(
                self.entry_qtd.get()), self.entry_local.get()
            if not all([sku, nome, local]): raise ValueError

            nova_linha = {
                'SKU': str(sku), 'Produto': nome, 'Qtd': qtd,
                'Endereco': local.upper(), 'Data': pd.Timestamp.now().strftime("%d/%m/%Y")
            }
            self.estoque = pd.concat([self.estoque, pd.DataFrame([nova_linha])], ignore_index=True)
            self.salvar_e_atualizar()
            self.limpar_campos()
        except:
            messagebox.showerror("Erro", "Preencha todos os campos corretamente (Qtd deve ser número)!")

    def registrar_saida(self):
        sku = self.entry_sku.get()
        if sku in self.estoque['SKU'].astype(str).values:
            idx = self.estoque[self.estoque['SKU'].astype(str) == sku].index[0]
            self.estoque.drop(idx, inplace=True)
            self.salvar_e_atualizar()
            messagebox.showinfo("Sucesso", f"Item SKU {sku} removido do estoque.")
        else:
            messagebox.showwarning("Erro", "SKU não encontrado para saída.")

    def salvar_e_atualizar(self):
        try:
            self.estoque.to_excel(self.arquivo_excel, index=False)
            self.mostrar_estoque()
        except:
            messagebox.showerror("Erro Crítico", "Feche o arquivo Excel antes de tentar salvar!")

    def filtrar_busca(self):
        termo = self.entry_busca.get().lower()
        res = self.estoque[
            (self.estoque['SKU'].astype(str).str.lower().str.contains(termo)) |
            (self.estoque['Produto'].str.lower().str.contains(termo)) |
            (self.estoque['Endereco'].str.lower().str.contains(termo))
            ]
        self.atualizar_visualizacao(res, f"🔍 RESULTADOS PARA: '{termo}'")

    def mostrar_estoque(self):
        self.atualizar_visualizacao(self.estoque, "📋 INVENTÁRIO COMPLETO")
        criticos = self.estoque[self.estoque['Qtd'] < self.limite_minimo].shape[0]
        if criticos > 0:
            self.lbl_alerta.configure(text=f"⚠️ ALERTA: {criticos} itens abaixo do nível de segurança!",
                                      text_color="#E74C3C")
        else:
            self.lbl_alerta.configure(text="✅ Todos os níveis de estoque estão estáveis.", text_color="#27AE60")

    def gerar_relatorio_reposicao(self):
        itens_baixos = self.estoque[self.estoque['Qtd'] < self.limite_minimo]
        if not itens_baixos.empty:
            itens_baixos.to_excel(self.arquivo_reposicao, index=False)
            messagebox.showinfo("Relatório", f"Lista de reposição gerada: {self.arquivo_reposicao}")
        else:
            messagebox.showinfo("Info", "Estoque OK. Sem necessidade de reposição.")

    def atualizar_visualizacao(self, df, titulo):
        self.textbox.delete("1.0", "end")
        self.textbox.insert("insert", f"{titulo}\n" + "=" * 85 + "\n\n")
        header = f"{'SKU':<12} | {'PRODUTO':<25} | {'QTD':<6} | {'LOCAL':<12} | {'DATA'}"
        self.textbox.insert("insert", header + "\n" + "-" * 85 + "\n")
        for _, row in df.iterrows():
            aviso = " [!]" if row['Qtd'] < self.limite_minimo else ""
            line = f"{str(row['SKU']):<12} | {str(row['Produto']):<25} | {str(row['Qtd']):<6} | {str(row['Endereco']):<12} | {row['Data']}{aviso}"
            self.textbox.insert("insert", line + "\n")

    def limpar_campos(self):
        for e in [self.entry_sku, self.entry_nome, self.entry_qtd, self.entry_local]: e.delete(0, 'end')


# --- INICIALIZAÇÃO DO PROGRAMA ---
if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    login_app = LoginWindow()
    login_app.mainloop()