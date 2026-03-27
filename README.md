# 📦 WMS Enterprise v12.8 - Warehouse Management System

O **WMS Enterprise** é uma solução robusta e moderna de Gerenciamento de Armazém desenvolvida em **Python**. O sistema foca na automação do recebimento via XML de NF-e, controle de estoque em tempo real, expedição de mercadorias e auditoria de processos.

## 🚀 Funcionalidades Principais

* **📥 Recebimento Inteligente:** Importação de arquivos XML de Notas Fiscais com conferência cega e detecção automática de divergências.
* **📦 Gestão de Estoque:** Controle total de SKUs, quantidades, endereçamentos e status (Disponível, Quarentena, Avariado).
* **🚚 Expedição Agilizada:** Rotina de baixa de estoque integrada com número de pedido de saída.
* **📊 Dashboard Dinâmico:** Visualização gráfica da ocupação e status do inventário utilizando Matplotlib.
* **📜 Auditoria Total:** Registro de Log (Histórico) de todas as movimentações (Entrada/Saída) com opção de exportação para **Excel**.
* **🔓 SKU Flexível:** Suporte para entrada de itens via XML ou cadastro manual livre.

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** [Python 3.11+](https://www.python.org/)
* **Interface Gráfica:** [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) (Interface moderna e Dark Mode)
* **Banco de Dados:** Microsoft Access (.accdb) via `pyodbc`
* **Processamento de Dados:** Pandas
* **Gráficos:** Matplotlib
* **Compilação:** PyInstaller
* **Usuario:** teste **Senha:** 1234

## 📋 Pré-requisitos

```bash
pip install customtkinter pyodbc pandas matplotlib openpyxl

🔧 Instalação e Execução
Clone o repositório: python main.py

📦 Como gerar o Executável (.exe)
Para transformar o script em um executável único para Windows:
pyinstaller --noconfirm --onefile --windowed --add-data "C:/Caminho/Para/customtkinter;customtkinter/" main(wms).py

🤝 Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir uma Issue ou enviar um Pull Request.

Faça um Fork do projeto
Crie uma branch para sua funcionalidade (git checkout -b feature/NovaFeature)
Commit suas mudanças (git commit -m 'Adicionando nova funcionalidade')
Push para a branch (git push origin feature/NovaFeature)
Abra um Pull Request
