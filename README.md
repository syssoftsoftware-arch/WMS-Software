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

## 📋 Pré-requisitos

Antes de rodar o projeto, você precisará instalar as dependências necessárias:

```bash
pip install customtkinter pyodbc pandas matplotlib openpyxl
