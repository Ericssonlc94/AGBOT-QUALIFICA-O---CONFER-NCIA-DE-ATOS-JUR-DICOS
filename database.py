import sqlite3
import os
import sys # Adicionado
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog

class Database:
    def __init__(self, db_file="municipios.db"):
        # Lógica para encontrar o DB tanto no modo de desenvolvimento quanto no executável
        if getattr(sys, 'frozen', False):
            # Se estiver rodando como um executável (congelado pelo PyInstaller)
            application_path = os.path.dirname(sys.executable)
        else:
            # Se estiver rodando como um script normal
            application_path = os.path.abspath(".")
        
        self.db_file = os.path.join(application_path, db_file)
        self.conn = sqlite3.connect(self.db_file)
        self.create_table()

    def create_table(self):
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS municipios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_municipio TEXT NOT NULL,
                ente TEXT NOT NULL,
                ano_exercicio TEXT,
                mes_referencia TEXT,
                cnpj TEXT,
                nome_prefeito TEXT,
                cpf TEXT
            )
        """)
        
        # Lógica de Migração de Esquema para adicionar a coluna faltante
        cursor.execute("PRAGMA table_info(municipios)")
        columns = [info[1] for info in cursor.fetchall()]

        if 'mes_referencia' not in columns:
            try:
                print("INFO: Atualizando esquema do banco de dados. Adicionando coluna 'mes_referencia'...")
                cursor.execute("ALTER TABLE municipios ADD COLUMN mes_referencia TEXT")
                print("INFO: Esquema do banco de dados atualizado com sucesso.")
            except Exception as e:
                print(f"ERRO: Falha ao atualizar o esquema do banco de dados: {e}")

        self.conn.commit()

    def adicionar_municipio(self):
        os.system('cls')
        print("+----------------------------------------------------------+")
        print("|                   ADICIONAR MUNICÍPIO                  |")
        print("+----------------------------------------------------------+")
        nome_municipio = input("Nome do Município: ")
        ente = input("Ente (Ex: PREFEITURA MUNICIPAL DE...): ")

        if not all([nome_municipio, ente]):
            print("\nErro: Nome do município e ente são obrigatórios.")
            return

        cursor = self.conn.cursor()
        # Insere valores vazios nos campos legados para satisfazer a constraint NOT NULL em DBs antigos
        try:
            cursor.execute("""
                INSERT INTO municipios (nome_municipio, ente, ano_exercicio, mes_referencia, cnpj, nome_prefeito, cpf)
                VALUES (?, ?, ?, ?, ?, ?, ?)""", 
                (nome_municipio, ente, "", "", "", "", ""))
        except sqlite3.IntegrityError as e:
            print(f"\nERRO: Ocorreu um erro de integridade no banco de dados. {e}")
            print("Isso pode acontecer se o nome do município já existir ou outra restrição foi violada.")
            return
            
        self.conn.commit()
        print("\nMunicípio adicionado com sucesso!")

    def add_municipio_from_gui(self, nome_municipio, ente):
        if not all([nome_municipio, ente]):
            raise ValueError("Nome do município e ente são obrigatórios.")

        cursor = self.conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO municipios (nome_municipio, ente, ano_exercicio, mes_referencia, cnpj, nome_prefeito, cpf)
                VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (nome_municipio, ente, "", "", "", "", ""))
        except sqlite3.IntegrityError as e:
            raise ValueError(f"Erro de integridade no banco de dados: {e}. O município pode já existir.")
            
        self.conn.commit()

    def consultar_municipios(self):
        os.system('cls')
        print("+----------------------------------------------------------+")
        print("|                 MUNICÍPIOS CADASTRADOS                 |")
        print("+----------------------------------------------------------+\n")
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, nome_municipio, ano_exercicio FROM municipios")
        municipios = cursor.fetchall()

        if not municipios:
            print("Nenhum município cadastrado.")
        else:
            for municipio in municipios:
                print(f"{municipio[0]}. {municipio[1]} ({municipio[2]})")
        
        print("\n1. Atualizar banco com txt ou excel")
        print("2. Voltar")
        choice = input("Escolha uma opção: ")
        if choice == '1':
            self.atualizar_banco()

    def editar_municipio(self):
        os.system('cls')
        print("+----------------------------------------------------------+")
        print("|                     EDITAR MUNICÍPIO                   |")
        print("+----------------------------------------------------------+")
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, nome_municipio FROM municipios")
        municipios = cursor.fetchall()

        if not municipios:
            print("Nenhum município cadastrado.")
            return

        for municipio in municipios:
            print(f"{municipio[0]}. {municipio[1]}")
        
        try:
            choice = int(input("Escolha o número do município para editar: "))
            
            cursor.execute("SELECT * FROM municipios WHERE id = ?", (choice,))
            municipio_data = cursor.fetchone()

            if municipio_data:
                print(f"\nEditando {municipio_data[1]}")

                nome_municipio = input(f"Nome do Município [{municipio_data[1]}]: ") or municipio_data[1]
                cnpj = input(f"CNPJ [{municipio_data[2]}]: ") or municipio_data[2]
                ente = input(f"Ente [{municipio_data[3]}]: ") or municipio_data[3]
                nome_prefeito = input(f"Nome do Prefeito/Presidente [{municipio_data[4]}]: ") or municipio_data[4]
                cpf = input(f"CPF [{municipio_data[5]}]: ") or municipio_data[5]
                ano_exercicio = input(f"Ano de Exercício [{municipio_data[6]}]: ") or municipio_data[6]

                cursor.execute("""
                    UPDATE municipios
                    SET nome_municipio = ?, cnpj = ?, ente = ?, nome_prefeito = ?, cpf = ?, ano_exercicio = ?
                    WHERE id = ?""", 
                    (nome_municipio, cnpj, ente, nome_prefeito, cpf, ano_exercicio, choice))
                self.conn.commit()
                print("\nMunicípio atualizado com sucesso!")
                input("Pressione Enter para voltar ao menu principal...")
            else:
                print("Escolha inválida.")
        except ValueError:
            print("Entrada inválida.")

    def excluir_municipio(self):
        os.system('cls')
        print("+----------------------------------------------------------+")
        print("|                   EXCLUIR MUNICÍPIO                    |")
        print("+----------------------------------------------------------+")
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, nome_municipio FROM municipios")
        municipios = cursor.fetchall()

        if not municipios:
            print("Nenhum município cadastrado.")
            return

        for municipio in municipios:
            print(f"{municipio[0]}. {municipio[1]}")
        
        try:
            choice = int(input("Escolha o número do município para excluir: "))
            
            cursor.execute("SELECT nome_municipio FROM municipios WHERE id = ?", (choice,))
            municipio_nome = cursor.fetchone()

            if municipio_nome:
                confirm = input(f"Tem certeza que deseja excluir {municipio_nome[0]}? (s/n): ")
                if confirm.lower() == 's':
                    cursor.execute("DELETE FROM municipios WHERE id = ?", (choice,))
                    self.conn.commit()
                    print("\nMunicípio excluído com sucesso!")
                    input("Pressione Enter para voltar ao menu principal...")
            else:
                print("Escolha inválida.")
        except ValueError:
            print("Entrada inválida.")

    def atualizar_banco(self):
        os.system('cls')
        print("+----------------------------------------------------------+")
        print("|               ATUALIZAR BANCO DE DADOS                 |")
        print("+----------------------------------------------------------+\n")
        file_path = self._ask_file_path("Selecione o arquivo .txt ou .xlsx para atualizar o banco", [("Text files", "*.txt"), ("Excel files", "*.xlsx")])
        if not file_path: return

        try:
            cursor = self.conn.cursor()
            if file_path.endswith('.xlsx'):
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False
                workbook = excel.Workbooks.Open(file_path)
                sheet = workbook.ActiveSheet
                
                for r_idx in range(2, sheet.UsedRange.Rows.Count + 1):
                    row = [sheet.Cells(r_idx, c_idx).Value for c_idx in range(1, 7)]
                    cursor.execute("""
                        INSERT INTO municipios (nome_municipio, cnpj, ente, nome_prefeito, cpf, ano_exercicio)
                        VALUES (?, ?, ?, ?, ?, ?)""", 
                        (row[0], row[1], row[2], row[3], row[4], row[5]))
                workbook.Close()
                excel.Quit()
            elif file_path.endswith('.txt'):
                with open(file_path, 'r') as f:
                    next(f) # Skip header
                    for line in f:
                        row = line.strip().split('\t')
                        cursor.execute("""
                            INSERT INTO municipios (nome_municipio, cnpj, ente, nome_prefeito, cpf, ano_exercicio)
                            VALUES (?, ?, ?, ?, ?, ?)""", 
                            (row[0], row[1], row[2], row[3], row[4], row[5]))

            self.conn.commit()
            print("Banco de dados atualizado com sucesso!")
            input("Pressione Enter para voltar ao menu principal...")

        except Exception as e:
            print(f"Ocorreu um erro ao ler o arquivo: {e}")
            
    def _ask_file_path(self, title, filetypes):
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return file_path
