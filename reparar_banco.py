import pyodbc
import os


def reparar_banco():
    db_path = os.path.abspath("wms_db.accdb")
    if not os.path.exists(db_path):
        print("❌ Erro: Arquivo wms_db.accdb não encontrado na pasta!")
        return

    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';'

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        print("🔗 Conectado ao banco. Iniciando auditoria de tabelas...")

        # 1. TABELA ESTOQUE
        try:
            cursor.execute("""
                           CREATE TABLE Estoque
                           (
                               SKU      TEXT(50),
                               Produto  TEXT(255),
                               Qtd      LONG,
                               Endereco TEXT(50),
                               Status   TEXT(50),
                               Lote     TEXT(50),
                               Validade TEXT(20)
                           )
                           """)
            print("✅ Tabela 'Estoque' criada.")
        except:
            print("ℹ️ Tabela 'Estoque' já existe.")

        # 2. TABELA HISTORICO
        try:
            cursor.execute("""
                           CREATE TABLE Historico
                           (
                               ID             AUTOINCREMENT PRIMARY KEY,
                               SKU            TEXT(50),
                               Tipo_Mov       TEXT(50),
                               Quantidade     LONG,
                               Origem_Destino TEXT(100),
                               Usuario        TEXT(50),
                               Data_Hora      TEXT(50)
                           )
                           """)
            print("✅ Tabela 'Historico' criada.")
        except:
            print("ℹ️ Tabela 'Historico' já existe.")

        # 3. TABELA MAPA_ARMAZEM (Essencial para a v12.5)
        try:
            cursor.execute("""
                           CREATE TABLE Mapa_Armazem
                           (
                               Endereco        TEXT(50) PRIMARY KEY,
                               Curva_Alvo      TEXT(1),
                               Status          TEXT(20),
                               Distancia_Docas LONG
                           )
                           """)
            # Inserindo endereços de teste se a tabela for nova
            cursor.execute("INSERT INTO Mapa_Armazem VALUES ('PR-01-A', 'A', 'Vazio', 10)")
            cursor.execute("INSERT INTO Mapa_Armazem VALUES ('PR-02-B', 'B', 'Vazio', 20)")
            cursor.execute("INSERT INTO Mapa_Armazem VALUES ('PR-03-C', 'C', 'Vazio', 30)")
            print("✅ Tabela 'Mapa_Armazem' criada com dados de teste.")
        except:
            print("ℹ️ Tabela 'Mapa_Armazem' já existe.")

        # 4. TABELA USUARIOS
        try:
            cursor.execute("CREATE TABLE Usuarios (Login TEXT(50), Senha TEXT(50), Perfil TEXT(20))")
            cursor.execute("INSERT INTO Usuarios VALUES ('admin', '123', 'ADM')")
            print("✅ Tabela 'Usuarios' criada. Login padrão: admin / 123")
        except:
            print("ℹ️ Tabela 'Usuarios' já existe.")

        conn.commit()
        print("\n🚀 BANCO PRONTO! Agora os dados devem aparecer no WMS.")

    except Exception as e:
        print(f"❌ Erro crítico: {e}")
    finally:
        if 'conn' in locals(): conn.close()


if __name__ == "__main__":
    reparar_banco()