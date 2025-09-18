import os
import shutil
import subprocess # Adicionado para fechar processos
import traceback
import win32com.client as win32
from database import Database
from utils import ask_file_path

# Definindo constantes do Excel manualmente para evitar problemas de importação
XL_FORMULAS = -4123
XL_PART = 2
XL_BY_ROWS = 1
XL_PREVIOUS = 2

class PreenchedorQualificacao:
    def __init__(self):
        """Inicializa os componentes necessários."""
        self.db = Database()
        print("PreenchedorQualificacao inicializado.")

    def run(self):
        """Exibe o menu principal e gerencia a navegação."""
        while True:
            os.system('cls')
            print("+------------------------------------------------------+")
            print("|    PREENCHEDOR DE QUALIFICAÇÃO - CONTRATOS/LICITAÇÕES|")
            print("+------------------------------------------------------+")
            print("| 1. Iniciar Preenchimento                             |")
            print("| 2. Adicionar Município                               |")
            print("| 3. Consultar Municípios                              |")
            print("| 4. Remover Município                                 |")
            print("| 5. Sair                                              |")
            print("+------------------------------------------------------+")
            
            choice = input("Escolha uma opção: ")

            if choice == '1':
                self.iniciar_preenchimento()
            elif choice == '2':
                self.gerenciar_municipio(self.db.adicionar_municipio, "ADICIONAR MUNICÍPIO")
            elif choice == '3':
                self.gerenciar_municipio(self.db.consultar_municipios, "CONSULTAR MUNICÍPIOS")
            elif choice == '4':
                self.gerenciar_municipio(self.db.excluir_municipio, "REMOVER MUNICÍPIO")
            elif choice == '5':
                break
            else:
                print("Opção inválida. Tente novamente.")
                input("Pressione Enter para continuar...")

    def gerenciar_municipio(self, funcao_db, titulo):
        """Função genérica para encapsular as operações de DB com limpeza de tela."""
        os.system('cls')
        print(f"""+-------------------------+
| {titulo.center(23)} |
+-------------------------+""")
        try:
            funcao_db()
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
        input("\nPressione Enter para voltar ao menu principal...")

    def execute_all(self, file_paths, municipio_nome, exercicio, mes, status_callback, output_dir=None):
        """Lógica de processamento principal, adaptada para ser chamada pela GUI ou CLI."""
        excel = None
        workbook = None
        try:
            status_callback("Fechando instâncias do Excel...", 0.05)
            try:
                subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], check=False, capture_output=True)
            except FileNotFoundError:
                status_callback("AVISO: 'taskkill' não encontrado. Não foi possível fechar o Excel.", 0.05)

            status_callback("Buscando dados do município...", 0.1)
            cursor = self.db.conn.cursor()
            cursor.execute("SELECT * FROM municipios WHERE nome_municipio = ?", (municipio_nome,))
            municipio_data = cursor.fetchone()
            if not municipio_data:
                raise ValueError(f"Município '{municipio_nome}' não encontrado no banco de dados.")

            status_callback("Preparando ambiente...", 0.15)
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False

            ente = municipio_data[2]
            modelo_path = file_paths['modelo']
            
            # Define o diretório de saída: usa o diretório passado pela GUI ou, como fallback, o diretório do arquivo de modelo.
            dir_name = output_dir if output_dir else os.path.dirname(modelo_path)
            new_base_name = f"Atos Jurídicos - Fator X Siga - {ente} - {mes}-{exercicio}.xlsx"
            new_path = os.path.join(dir_name, new_base_name)
            
            if os.path.exists(new_path):
                os.remove(new_path)

            shutil.copy(modelo_path, new_path)
            status_callback(f"Arquivo de saída criado: {new_base_name}", 0.2)

            workbook = excel.Workbooks.Open(new_path)

            # Processar Licitações
            status_callback("Processando Licitações...", 0.3)
            dados_licitacoes_siga = self._extrair_dados_licitacoes_html(excel, file_paths['licitacoes_html'])
            dados_listagem_fator = self._extrair_dados_listagem(excel, file_paths['listagem_licitacoes'])
            dados_licitacoes_unificados = self._unificar_e_ordenar_dados(dados_licitacoes_siga, dados_listagem_fator)
            self._preencher_licitacoes_modelo(workbook, dados_licitacoes_unificados)

            # Processar Dispensas e Inexigibilidades
            status_callback("Processando Dispensas e Inexigibilidades...", 0.5)
            dados_dispensas_siga = self._extrair_dados_dispensas_html(excel, file_paths['dispensas_html'])
            dados_listagem_dispensas = self._extrair_dados_listagem_dispensas_excel(excel, file_paths['listagem_dispensas'])
            dados_dispensas_unificados = self._unificar_e_ordenar_dados_dispensas(dados_dispensas_siga, dados_listagem_dispensas)
            self._preencher_dispensas_inex_modelo(workbook, dados_dispensas_unificados)

            # Processar Contratos
            status_callback("Processando Contratos...", 0.7)
            dados_contratos_siga = self._extrair_dados_contratos_html(excel, file_paths['contratos_html'])
            dados_listagem_contratos = self._extrair_dados_listagem_contratos_excel(excel, file_paths['listagem_contratos'])
            dados_contratos_unificados = self._unificar_e_ordenar_dados_contratos(dados_contratos_siga, dados_listagem_contratos)
            self._preencher_contratos_modelo(workbook, dados_contratos_unificados)

            status_callback("Salvando arquivo final...", 0.9)
            workbook.Save()
            status_callback("Processo concluído com sucesso!", 1.0)

        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel:
                excel.Quit()

    def iniciar_preenchimento(self):
        # Este método agora serve principalmente para o modo CLI.
        # A lógica principal foi movida para execute_all
        os.system('cls')
        print("+------------------------------------------------------+")
        print("|              INICIAR PREENCHIMENTO (CLI)             |")
        print("+------------------------------------------------------+")
        
        try:
            municipio_data = self._selecionar_municipio()
            if not municipio_data: return

            exercicio = input("Digite o Ano de Exercício (Ex: 2025): ")
            mes = input("Digite o Mês de Referência (Ex: Julho): ")
            if not all([exercicio, mes]):
                print("Exercício e Mês são obrigatórios.")
                return

            file_paths = {}
            print("\n--- Seleção de Arquivos ---")
            file_paths["modelo"] = ask_file_path("Selecione o modelo 'Conferencia mensal...", [("Excel files", "*.xlsx")])
            file_paths["licitacoes_html"] = ask_file_path("Selecione o HTML 'Licitações SIGA'", [("HTML files", "*.htm;*.html")])
            file_paths["listagem_licitacoes"] = ask_file_path("Selecione o Excel 'Listagem de Licitação'", [("Excel files", "*.xlsx;*.xls")])
            file_paths["dispensas_html"] = ask_file_path("Selecione o HTML 'Dispensas e Inex SIGA'", [("HTML files", "*.htm;*.html")])
            file_paths["listagem_dispensas"] = ask_file_path("Selecione o Excel 'Listagem de Dispensas e Inex'", [("Excel files", "*.xlsx;*.xls")])
            file_paths["contratos_html"] = ask_file_path("Selecione o HTML 'Contratos SIGA'", [("HTML files", "*.htm;*.html")])
            file_paths["listagem_contratos"] = ask_file_path("Selecione o Excel 'Listagem de Contratos'", [("Excel files", "*.xlsx;*.xls")])

            if not all(file_paths.values()):
                print("Todos os arquivos devem ser selecionados.")
                return

            # Define o diretório de saída para o modo CLI
            output_dir_cli = os.path.dirname(file_paths["licitacoes_html"])

            # Callback simples para printar o status no console
            def cli_status_callback(message, progress):
                print(f"[{int(progress*100)}%] {message}")

            self.execute_all(file_paths, municipio_data[1], exercicio, mes, cli_status_callback, output_dir=output_dir_cli)
            print("\nProcesso concluído com sucesso!")

        except Exception as e:
            print(f"\nOcorreu um erro inesperado no processo: {e}")
            traceback.print_exc()
        finally:
            input("\nPressione Enter para voltar ao menu principal...")

    def _selecionar_municipio(self):
        """Lista e permite ao usuário selecionar um município."""
        cursor = self.db.conn.cursor()
        cursor.execute("SELECT id, nome_municipio FROM municipios")
        municipios = cursor.fetchall()

        if not municipios:
            print("Nenhum município cadastrado.")
            return None

        print("Municípios disponíveis:")
        for mid, nome in municipios:
            print(f"{mid}. {nome}")
        
        try:
            choice = int(input("Escolha o número do município: "))
            cursor.execute("SELECT * FROM municipios WHERE id = ?", (choice,))
            municipio_data = cursor.fetchone()
            if not municipio_data:
                print("Escolha inválida.")
                return None
            print(f"Município selecionado: {municipio_data[1]}")
            return municipio_data
        except ValueError:
            print("Entrada inválida.")
            return None

    

    def _extrair_dados_licitacoes_html(self, excel, html_path):
        """Extrai dados do arquivo HTML de licitações do SIGA."""
        print(f"Extraindo dados de: {os.path.basename(html_path)}")
        wb = None
        try:
            wb = excel.Workbooks.Open(html_path)
            sheet = wb.Worksheets(1)
            last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row

            processos = []
            valores = []

            # Assumindo colunas: Processo=2 (B), Valor Estimado=6 (F)
            for i in range(2, last_row + 1):
                processo = sheet.Cells(i, 2).Value
                valor_str = str(sheet.Cells(i, 6).Value)

                if processo:
                    processos.append(str(processo))
                    valores.append(self._formatar_valor(valor_str))

            print(f"Encontrados {len(processos)} registros no HTML de Licitações.")
            return {"processos": processos, "valores": valores}
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def _extrair_dados_listagem(self, excel, listagem_path):
        """Extrai dados do arquivo Excel de listagem de licitações."""
        print(f"Extraindo dados de: {os.path.basename(listagem_path)}")
        wb = None
        try:
            wb = excel.Workbooks.Open(listagem_path)
            sheet = wb.Worksheets(1)
            sheet.UsedRange.UnMerge()

            # Find the last row with any content in the sheet
            last_cell = sheet.Cells.Find(What="*", After=sheet.Cells(1, 1), LookIn=XL_FORMULAS,
                                         LookAt=XL_PART, SearchOrder=XL_BY_ROWS,
                                         SearchDirection=XL_PREVIOUS, MatchCase=False)
            if last_cell:
                last_row = last_cell.Row
            else:
                last_row = 1 # If no data found, assume only header or empty
            print(f"DEBUG: Última linha encontrada com Find: {last_row}")
            
            numeros = []
            valores = []

            # Assumindo colunas: Número=2 (B), Valor Estimado=21 (U)
            # Começando da linha 9
            for i in range(9, last_row + 1):
                # Verifica se a linha contém "Total de Registros" e para a leitura
                # Vou verificar na coluna B (onde está o número da licitação)
                cell_b_value = str(sheet.Cells(i, 2).Value).strip()
                if "total de registros" in cell_b_value.lower():
                    print(f"DEBUG: Linha {i}: 'Total de Registros' encontrado. Parando a leitura.")
                    break

                numero_val = sheet.Cells(i, 2).Value # Column B
                valor_str = str(sheet.Cells(i, 21).Value) # Column U
                
                print(f"DEBUG: Linha {i}: Coluna B (Número) = '{numero_val}' (Tipo: {type(numero_val)}), Coluna U (Valor Estimado) = '{valor_str}'")
                
                if numero_val is not None and str(numero_val).strip() != "":
                    numeros.append(str(numero_val).strip())
                    valores.append(self._formatar_valor(valor_str))
                else:
                    print(f"DEBUG: Linha {i}: Valor da coluna B ignorado (vazio ou None).\n")


            print(f"Encontrados {len(numeros)} registros na Listagem de Licitação.")
            return {"numeros": numeros, "valores": valores}
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def _unificar_e_ordenar_dados(self, dados_siga, dados_fator):
        """Unifica e ordena os dados do SIGA e FATOR por número/processo."""
        print("Unificando e ordenando dados...")
        unified_data = {}

        # Processa dados do SIGA (HTML)
        for i, processo_siga in enumerate(dados_siga["processos"]):
            valor_siga = dados_siga["valores"][i]
            unified_data[str(processo_siga)] = { # Ensure key is string
                "siga_processo": str(processo_siga), # Ensure value is string
                "siga_valor": valor_siga,
                "fator_numero": "",
                "fator_valor": 0.0
            }

        # Processa dados do FATOR (Excel)
        for i, numero_fator in enumerate(dados_fator["numeros"]):
            valor_fator = dados_fator["valores"][i]
            if str(numero_fator) in unified_data: # Ensure key is string for lookup
                # Atualiza entrada existente
                unified_data[str(numero_fator)]["fator_numero"] = str(numero_fator) # Ensure value is string
                unified_data[str(numero_fator)]["fator_valor"] = valor_fator
            else:
                # Adiciona nova entrada
                unified_data[str(numero_fator)] = { # Ensure key is string
                    "siga_processo": "",
                    "siga_valor": 0.0,
                    "fator_numero": str(numero_fator), # Ensure value is string
                    "fator_valor": valor_fator
                }
        
        # Converte para lista e ordena pelo número/processo
        sorted_records = sorted(unified_data.values(), key=lambda x: str(x.get("siga_processo") or x.get("fator_numero")))
        print(f"Dados unificados e ordenados: {len(sorted_records)} registros.")
        return sorted_records

    def _preencher_licitacoes_modelo(self, workbook, unified_records):
        """Preenche a aba 'LICITAÇÕES' da planilha com os dados unificados."""
        print("Iniciando preenchimento da aba 'LICITAÇÕES'...")
        try:
            sheet = workbook.Worksheets("LICITAÇÕES")

            # Limpa áreas de colagem (novas e antigas)
            sheet.Range("A4:B1000").ClearContents()
            sheet.Range("H4:I1000").ClearContents()
            sheet.Range("K4:L1000").ClearContents()
            print("Áreas antigas limpas na aba 'LICITAÇÕES'.")

            if not unified_records:
                print("Nenhum dado de licitação para preencher.")
                return

            # Prepara os dados para colagem em massa
            fator_numeros = [[rec["fator_numero"]] for rec in unified_records]
            siga_processos = [[rec["siga_processo"]] for rec in unified_records]
            fator_valores = [[rec["fator_valor"]] for rec in unified_records]
            siga_valores = [[rec["siga_valor"]] for rec in unified_records]

            # Cola os dados em massa
            start_row = 4
            end_row = start_row + len(unified_records) - 1
            sheet.Range(f"A{start_row}:A{end_row}").Value = fator_numeros
            sheet.Range(f"B{start_row}:B{end_row}").Value = siga_processos
            sheet.Range(f"H{start_row}:H{end_row}").Value = fator_valores # Coluna H para Fator Valor
            sheet.Range(f"I{start_row}:I{end_row}").Value = siga_valores # Coluna I para SIGA Valor
            
            print(f"{len(unified_records)} registros colados com sucesso na aba 'LICITAÇÕES'.")

        except Exception as e:
            print(f"ERRO ao preencher a aba 'LICITAÇÕES': {e}")

    def _extrair_dados_dispensas_html(self, excel, html_path):
        """Extrai dados do arquivo HTML de dispensas e inexigibilidades do SIGA."""
        print(f"Extraindo dados de: {os.path.basename(html_path)}")
        wb = None
        try:
            wb = excel.Workbooks.Open(html_path)
            sheet = wb.Worksheets(1)
            
            last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
            
            processos = []
            valores = []

            # Assumindo colunas: Processo=2 (B), Valor Total=6 (F)
            for i in range(2, last_row + 1):
                processo = sheet.Cells(i, 2).Value
                valor_str = str(sheet.Cells(i, 6).Value)

                if processo:
                    processos.append(processo)
                    valores.append(self._formatar_valor(valor_str))

            print(f"Encontrados {len(processos)} registros no HTML de Dispensas/Inexigibilidades.")
            return {"processos": processos, "valores": valores}
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def _extrair_dados_listagem_dispensas_excel(self, excel, listagem_path):
        """Extrai dados do arquivo Excel de listagem de dispensas e inexigibilidades."""
        print(f"Extraindo dados de: {os.path.basename(listagem_path)}")
        wb = None
        try:
            wb = excel.Workbooks.Open(listagem_path)
            sheet = wb.Worksheets(1)
            sheet.UsedRange.UnMerge()

            last_cell = sheet.Cells.Find(What="*", After=sheet.Cells(1, 1), LookIn=XL_FORMULAS,
                                         LookAt=XL_PART, SearchOrder=XL_BY_ROWS,
                                         SearchDirection=XL_PREVIOUS, MatchCase=False)
            if last_cell:
                last_row = last_cell.Row
            else:
                last_row = 1
            print(f"DEBUG: Última linha encontrada com Find: {last_row}")
            
            numeros = []
            valores = []

            for i in range(9, last_row + 1):
                cell_b_value = str(sheet.Cells(i, 2).Value).strip()
                if "total de registros" in cell_b_value.lower():
                    print(f"DEBUG: Linha {i}: 'Total de Registros' encontrado. Parando a leitura.")
                    break

                numero_val = sheet.Cells(i, 2).Value
                valor_str = str(sheet.Cells(i, 21).Value)
                
                print(f"DEBUG: Linha {i}: Coluna B (Número) = '{numero_val}' (Tipo: {type(numero_val)}), Coluna U (Valor Estimado) = '{valor_str}'")
                
                if numero_val is not None and str(numero_val).strip() != "":
                    numeros.append(str(numero_val).strip())
                    valores.append(self._formatar_valor(valor_str))
                else:
                    print(f"DEBUG: Linha {i}: Valor da coluna B ignorado (vazio ou None).\n")

            print(f"Encontrados {len(numeros)} registros na Listagem de Dispensa/Inexigibilidade.")
            return {"numeros": numeros, "valores": valores}
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def _unificar_e_ordenar_dados_dispensas(self, dados_siga, dados_fator):
        """Unifica e ordena os dados de Dispensas/Inexigibilidades por número/processo."""
        print("Unificando e ordenando dados de Dispensas/Inexigibilidades...")
        unified_data = {}

        # Processa dados do SIGA (HTML)
        for i, processo_siga in enumerate(dados_siga["processos"]):
            valor_siga = dados_siga["valores"][i]
            key = str(processo_siga)
            unified_data[key] = {
                "siga_processo": key,
                "siga_valor": valor_siga,
                "fator_numero": "",
                "fator_valor": 0.0
            }

        # Processa dados do FATOR (Excel)
        for i, numero_fator in enumerate(dados_fator["numeros"]):
            valor_fator = dados_fator["valores"][i]
            key = str(numero_fator)
            if key in unified_data:
                # Atualiza entrada existente
                unified_data[key]["fator_numero"] = key
                unified_data[key]["fator_valor"] = valor_fator
            else:
                # Adiciona nova entrada
                unified_data[key] = {
                    "siga_processo": "",
                    "siga_valor": 0.0,
                    "fator_numero": key,
                    "fator_valor": valor_fator
                }
        
        # Converte para lista e ordena pelo número/processo
        sorted_records = sorted(unified_data.values(), key=lambda x: str(x.get("siga_processo") or x.get("fator_numero")))
        print(f"Dados unificados e ordenados de Dispensas/Inexigibilidades: {len(sorted_records)} registros.")
        return sorted_records

    def _preencher_dispensas_inex_modelo(self, workbook, unified_records):
        """Preenche a aba 'DISPENSAS-INEX' da planilha com os dados unificados."""
        print("Iniciando preenchimento da aba 'DISPENSAS-INEX'...")
        try:
            sheet = workbook.Worksheets("DISPENSAS-INEX")

            # Limpa áreas de colagem (novas e antigas)
            sheet.Range("A4:B1000").ClearContents()
            sheet.Range("H4:I1000").ClearContents()
            sheet.Range("K4:L1000").ClearContents()
            print("Áreas antigas limpas na aba 'DISPENSAS-INEX'.")

            if not unified_records:
                print("Nenhum dado de dispensa/inexigibilidade para preencher.")
                return

            # Prepara os dados para colagem em massa
            fator_numeros = [[rec["fator_numero"]] for rec in unified_records]
            siga_processos = [[rec["siga_processo"]] for rec in unified_records]
            fator_valores = [[rec["fator_valor"]] for rec in unified_records]
            siga_valores = [[rec["siga_valor"]] for rec in unified_records]

            # Cola os dados em massa
            start_row = 4
            end_row = start_row + len(unified_records) - 1
            sheet.Range(f"A{start_row}:A{end_row}").Value = fator_numeros
            sheet.Range(f"B{start_row}:B{end_row}").Value = siga_processos
            sheet.Range(f"H{start_row}:H{end_row}").Value = fator_valores # Coluna H para Fator Valor
            sheet.Range(f"I{start_row}:I{end_row}").Value = siga_valores # Coluna I para SIGA Valor

            print(f"{len(unified_records)} registros colados com sucesso na aba 'DISPENSAS-INEX'.")

        except Exception as e:
            print(f"ERRO ao preencher a aba 'DISPENSAS-INEX': {e}")

    def _extrair_dados_contratos_html(self, excel, html_path):
        """Extrai dados do arquivo HTML de contratos do SIGA."""
        print(f"Extraindo dados de: {os.path.basename(html_path)}")
        wb = None
        try:
            wb = excel.Workbooks.Open(html_path)
            sheet = wb.Worksheets(1)
            last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row

            processos = []
            valores = []

            # Assumindo colunas: Contrato=2 (B), Valor=7 (G)
            for i in range(2, last_row + 1):
                processo = sheet.Cells(i, 2).Value
                valor_str = str(sheet.Cells(i, 7).Value)

                if processo:
                    processos.append(str(processo))
                    valores.append(self._formatar_valor(valor_str))

            print(f"Encontrados {len(processos)} registros no HTML de Contratos.")
            return {"processos": processos, "valores": valores}
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def _extrair_dados_listagem_contratos_excel(self, excel, listagem_path):
        """Extrai dados do arquivo Excel de listagem de contratos."""
        print(f"Extraindo dados de: {os.path.basename(listagem_path)}")
        wb = None
        try:
            wb = excel.Workbooks.Open(listagem_path)
            sheet = wb.Worksheets(1)
            sheet.UsedRange.UnMerge()

            last_cell = sheet.Cells.Find(What="*", After=sheet.Cells(1, 1), LookIn=XL_FORMULAS,
                                         LookAt=XL_PART, SearchOrder=XL_BY_ROWS,
                                         SearchDirection=XL_PREVIOUS, MatchCase=False)
            if last_cell:
                last_row = last_cell.Row
            else:
                last_row = 1
            print(f"DEBUG: Última linha encontrada com Find: {last_row}")
            
            numeros = []
            valores = []

            for i in range(9, last_row + 1):
                cell_b_value = str(sheet.Cells(i, 2).Value).strip()
                if "total de registros" in cell_b_value.lower():
                    print(f"DEBUG: Linha {i}: 'Total de Registros' encontrado. Parando a leitura.")
                    break

                numero_val = sheet.Cells(i, 2).Value
                valor_str = str(sheet.Cells(i, 21).Value) # Assumindo coluna U (21)
                
                print(f"DEBUG: Linha {i}: Coluna B (Número) = '{numero_val}' (Tipo: {type(numero_val)}), Coluna U (Valor Estimado) = '{valor_str}'")
                
                if numero_val is not None and str(numero_val).strip() != "":
                    numeros.append(str(numero_val).strip())
                    valores.append(self._formatar_valor(valor_str))
                else:
                    print(f"DEBUG: Linha {i}: Valor da coluna B ignorado (vazio ou None).\n")

            print(f"Encontrados {len(numeros)} registros na Listagem de Contratos.")
            return {"numeros": numeros, "valores": valores}
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def _unificar_e_ordenar_dados_contratos(self, dados_siga, dados_fator):
        """Unifica e ordena os dados de Contratos por número/processo."""
        print("Unificando e ordenando dados de Contratos...")
        unified_data = {}

        # Processa dados do SIGA (HTML)
        for i, processo_siga in enumerate(dados_siga["processos"]):
            valor_siga = dados_siga["valores"][i]
            key = str(processo_siga)
            unified_data[key] = {
                "siga_processo": key,
                "siga_valor": valor_siga,
                "fator_numero": "",
                "fator_valor": 0.0
            }

        # Processa dados do FATOR (Excel)
        for i, numero_fator in enumerate(dados_fator["numeros"]):
            valor_fator = dados_fator["valores"][i]
            key = str(numero_fator)
            if key in unified_data:
                unified_data[key]["fator_numero"] = key
                unified_data[key]["fator_valor"] = valor_fator
            else:
                unified_data[key] = {
                    "siga_processo": "",
                    "siga_valor": 0.0,
                    "fator_numero": key,
                    "fator_valor": valor_fator
                }
        
        sorted_records = sorted(unified_data.values(), key=lambda x: str(x.get("siga_processo") or x.get("fator_numero")))
        print(f"Dados unificados e ordenados de Contratos: {len(sorted_records)} registros.")
        return sorted_records

    def _preencher_contratos_modelo(self, workbook, unified_records):
        """Preenche a aba 'CONTRATOS' da planilha com os dados unificados."""
        print("Iniciando preenchimento da aba 'CONTRATOS'...")
        try:
            sheet = workbook.Worksheets("CONTRATOS")

            # Limpa áreas de colagem
            sheet.Range("A4:B1000").ClearContents()
            sheet.Range("H4:I1000").ClearContents()
            sheet.Range("K4:L1000").ClearContents() # Also clear old cols just in case
            print("Áreas antigas limpas na aba 'CONTRATOS'.")

            if not unified_records:
                print("Nenhum dado de contrato para preencher.")
                return

            # Prepara os dados para colagem em massa
            fator_numeros = [[rec["fator_numero"]] for rec in unified_records]
            siga_processos = [[rec["siga_processo"]] for rec in unified_records]
            fator_valores = [[rec["fator_valor"]] for rec in unified_records]
            siga_valores = [[rec["siga_valor"]] for rec in unified_records]

            # Cola os dados em massa
            start_row = 4
            end_row = start_row + len(unified_records) - 1
            sheet.Range(f"A{start_row}:A{end_row}").Value = fator_numeros
            sheet.Range(f"B{start_row}:B{end_row}").Value = siga_processos
            sheet.Range(f"H{start_row}:H{end_row}").Value = fator_valores
            sheet.Range(f"I{start_row}:I{end_row}").Value = siga_valores

            print(f"{len(unified_records)} registros colados com sucesso na aba 'CONTRATOS'.")

        except Exception as e:
            print(f"ERRO ao preencher a aba 'CONTRATOS': {e}")

    def _formatar_valor(self, valor_str):
        """Converte uma string de valor para float."""
        if valor_str is None or valor_str.strip() == "" or valor_str == 'None':
            return 0.0
        try:
            # Remove R$, espaços, e troca vírgula por ponto
            return float(valor_str.replace("R$", "").strip().replace(".", "").replace(",", "."))
        except (ValueError, TypeError):
            return 0.0
            
    def _colar_coluna(self, sheet, celula_inicial, dados):
        """Cola uma lista de dados em uma coluna a partir de uma célula."""
        if not dados:
            return
        range_end_row = int(celula_inicial[1:]) + len(dados) - 1
        col = celula_inicial[0]
        target_range = sheet.Range(f"{col}{celula_inicial[1:]}:{col}{range_end_row}")
        target_range.Value = [[d] for d in dados]


if __name__ == "__main__":
    qualificador = PreenchedorQualificacao()
    qualificador.run()
