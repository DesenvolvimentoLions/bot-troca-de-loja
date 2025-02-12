import sqlite3

def read_db_file(db_path):
    try:
        # Conecta ao banco de dados
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Obtém a lista de tabelas no banco de dados
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # Itera sobre as tabelas e exibe seus conteúdos
        for table in tables:
            table_name = table[0]
            print(f"Conteúdo da tabela {table_name}:")
            cursor.execute(f"SELECT * FROM {table_name}")
            rows = cursor.fetchall()
            for row in rows:
                print(row)
            print("\n")  # Adiciona uma linha em branco entre as tabelas

    except sqlite3.Error as e:
        print(f"Erro ao acessar o banco de dados: {e}")
    finally:
        if conn:
            conn.close()

# Caminho para o arquivo .db
db_path = 'C:\Users\ryan.alves\Desktop\Bot carro\,.db'
read_db_file(db_path)