import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import threading
import traceback
import os # Importar os
import sys # Importar sys
from preenchedor_qualificacao import PreenchedorQualificacao
from database import Database

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Configurações de Cor e Aparência ---
ctk.set_appearance_mode("Dark")

LEAD_COLOR = "#4A4A4A"
GREEN_ACCENT = "#4CAF50"
GREEN_HOVER = "#388E3C"
WHITE_TEXT = "#FFFFFF"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuração da Janela Principal ---
        self.title("Preenchedor de Qualificação Fator X Siga")
        self.geometry("800x750")
        self.configure(fg_color=LEAD_COLOR)
        self.grid_columnconfigure(0, weight=1)

        self.file_paths = {}
        # Define all file types and their initial state
        all_file_keys = [
            "modelo", "licitacoes_html", "listagem_licitacoes",
            "dispensas_html", "listagem_dispensas",
            "contratos_html", "listagem_contratos"
        ]
        for key in all_file_keys:
            self.file_paths[key] = {"label": None, "path": None}

        self.db = Database() # Instancia o banco de dados aqui

        # Hardcode the template file path using the resource_path function
        template_file_name = "Modelo - Conferencia mensal FATOR X SIGA - Contratos, Licitações, Dispensas e Inex.xlsx"
        self.template_full_path = resource_path(template_file_name)
        self.file_paths["modelo"]["path"] = self.template_full_path # Update path for modelo

        # --- Definição de Fontes ---
        self.font_large_bold = ctk.CTkFont(family="Arial", size=24, weight="bold")
        self.font_medium_bold = ctk.CTkFont(family="Arial", size=14, weight="bold")
        self.font_normal = ctk.CTkFont(family="Arial", size=12)

        # --- Título ---
        title_label = ctk.CTkLabel(self, text="Preenchedor de Qualificação", font=self.font_large_bold)
        title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # --- Frame de Parâmetros ---
        params_frame = ctk.CTkFrame(self)
        params_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        params_frame.grid_columnconfigure(1, weight=1)

        self.create_parameter_widgets(params_frame)

        # --- Frame de Arquivos ---
        files_frame = ctk.CTkFrame(self)
        files_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        files_frame.grid_columnconfigure(1, weight=1)

        self.create_file_selection_widgets(files_frame)

        # --- Frame de Ação ---
        action_frame = ctk.CTkFrame(self)
        action_frame.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)

        self.create_action_widgets(action_frame)

    def create_parameter_widgets(self, parent_frame):
        # Município
        municipio_label = ctk.CTkLabel(parent_frame, text="Município:", anchor="w", font=self.font_normal)
        municipio_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.municipio_combobox = ctk.CTkComboBox(parent_frame, values=[], state="readonly", font=self.font_normal)
        self.municipio_combobox.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.populate_municipios_dropdown()

        # Ano
        ano_label = ctk.CTkLabel(parent_frame, text="Ano de Exercício:", anchor="w", font=self.font_normal)
        ano_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.ano_entry = ctk.CTkEntry(parent_frame, placeholder_text="Ex: 2025", font=self.font_normal)
        self.ano_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # Mês
        mes_label = ctk.CTkLabel(parent_frame, text="Mês de Referência:", anchor="w", font=self.font_normal)
        mes_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.mes_entry = ctk.CTkEntry(parent_frame, placeholder_text="Ex: Julho", font=self.font_normal)
        self.mes_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        # Botão Adicionar Município
        add_municipio_button = ctk.CTkButton(parent_frame, text="Adicionar Novo Município", command=self.add_municipio_window, font=self.font_normal)
        add_municipio_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

    def create_file_selection_widgets(self, parent_frame):
        file_types_config = {
            "modelo": {"text": "Modelo 'Conferencia mensal...' (Automático)", "filter": [("Excel files", "*.xlsx")]},
            "licitacoes_html": {"text": "HTML 'Licitações SIGA' (*.htm, *.html)", "filter": [("HTML files", "*.htm;*.html")]},
            "listagem_licitacoes": {"text": "Excel 'Listagem de Licitação' (*.xlsx, *.xls)", "filter": [("Excel files", "*.xlsx;*.xls")]},
            "dispensas_html": {"text": "HTML 'Dispensas e Inex SIGA' (*.htm, *.html)", "filter": [("HTML files", "*.htm;*.html")]},
            "listagem_dispensas": {"text": "Excel 'Listagem de Dispensas e Inex' (*.xlsx, *.xls)", "filter": [("Excel files", "*.xlsx;*.xls")]},
            "contratos_html": {"text": "HTML 'Contratos SIGA' (*.htm, *.html)", "filter": [("HTML files", "*.htm;*.html")]},
            "listagem_contratos": {"text": "Excel 'Listagem de Contratos' (*.xlsx, *.xls)", "filter": [("Excel files", "*.xlsx;*.xls")]},
        }
        
        for i, (key, config) in enumerate(file_types_config.items()):
            label = ctk.CTkLabel(parent_frame, text=config["text"], anchor="w", font=self.font_normal)
            label.grid(row=i, column=0, padx=10, pady=5, sticky="w")

            path_label = ctk.CTkLabel(parent_frame, text="Nenhum arquivo selecionado", text_color="gray", anchor="w", font=self.font_normal)
            path_label.grid(row=i, column=1, padx=10, pady=5, sticky="ew")

            # Store label reference for hardcoded path
            self.file_paths[key]["label"] = path_label

            if key == "modelo":
                # For hardcoded template, just update the label and disable button
                path_label.configure(text=os.path.basename(self.template_full_path), text_color=WHITE_TEXT)
                button = ctk.CTkButton(parent_frame, text="Selecionar...", state="disabled", font=self.font_normal)
            else:
                button = ctk.CTkButton(parent_frame, text="Selecionar...", command=lambda k=key, pl=path_label, ft=config["filter"], dt=config["text"]: self.select_file(k, pl, ft, dt), font=self.font_normal)
            
            button.grid(row=i, column=2, padx=10, pady=5)

    def create_action_widgets(self, parent_frame):
        self.run_button = ctk.CTkButton(parent_frame, text="Iniciar Preenchimento", font=self.font_medium_bold, fg_color=GREEN_ACCENT, hover_color=GREEN_HOVER, command=self.run_processing_thread)
        self.run_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(parent_frame, mode="determinate")
        self.progress_bar.set(0)
        self.progress_bar.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        self.status_label = ctk.CTkLabel(parent_frame, text="Aguardando início...", anchor="w", font=self.font_normal)
        self.status_label.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

    def select_file(self, key, path_label, filetypes_filter, dialog_title):
        file_path = filedialog.askopenfilename(filetypes=filetypes_filter, title=f"Selecione o arquivo: {dialog_title}") # Pass filetypes and title here
        if file_path:
            self.file_paths[key]["path"] = file_path
            path_label.configure(text=os.path.basename(file_path), text_color=WHITE_TEXT) # Mostra apenas o nome do arquivo

    def populate_municipios_dropdown(self):
        try:
            cursor = self.db.conn.cursor()
            cursor.execute("SELECT nome_municipio FROM municipios ORDER BY nome_municipio")
            municipios = [row[0] for row in cursor.fetchall()]
            self.municipio_combobox.configure(values=municipios)
            if municipios:
                self.municipio_combobox.set(municipios[0])
        except Exception as e:
            self.status_label.configure(text=f"Erro ao carregar municípios: {e}", text_color="orange")

    def add_municipio_window(self):
        window = ctk.CTkToplevel(self)
        window.title("Adicionar Município")
        window.geometry("300x200")
        window.transient(self) # Make it modal
        window.grab_set() # Grab all events until this window is destroyed

        ctk.CTkLabel(window, text="Nome do Município:", font=self.font_normal).pack(pady=5)
        self.new_municipio_name_entry = ctk.CTkEntry(window, font=self.font_normal)
        self.new_municipio_name_entry.pack(pady=5)

        ctk.CTkLabel(window, text="Ente (Ex: PM Barra da Estiva):", font=self.font_normal).pack(pady=5)
        self.new_municipio_ente_entry = ctk.CTkEntry(window, font=self.font_normal)
        self.new_municipio_ente_entry.pack(pady=5)

        ctk.CTkButton(window, text="Adicionar", command=lambda: self._submit_new_municipio(window), font=self.font_normal).pack(pady=10)

    def _submit_new_municipio(self, window):
        nome = self.new_municipio_name_entry.get()
        ente = self.new_municipio_ente_entry.get()

        if not nome or not ente:
            self.status_label.configure(text="Erro: Nome e Ente do município são obrigatórios.", text_color="orange")
            return

        try:
            self.db.add_municipio_from_gui(nome, ente)
            self.status_label.configure(text=f"Município '{nome}' adicionado com sucesso!", text_color="lightgreen")
            self.populate_municipios_dropdown() # Refresh dropdown
            window.destroy()
        except Exception as e:
            self.status_label.configure(text=f"Erro ao adicionar município: {e}", text_color="orange")

    def run_processing_thread(self):
        # Desabilitar o botão para prevenir cliques duplos
        self.run_button.configure(state="disabled")
        
        # Iniciar o processamento em uma nova thread para não congelar a GUI
        processing_thread = threading.Thread(target=self.processing_logic)
        processing_thread.start()

    def gui_status_update(self, message, progress):
        self.status_label.configure(text=message)
        self.progress_bar.set(progress)

    def processing_logic(self):
        """A lógica de processamento real que é executada em uma thread."""
        try:
            self.after(0, self.gui_status_update, "Validando entradas...", 0.05)

            # Coleta e validação dos caminhos dos arquivos
            paths = {key: data["path"] for key, data in self.file_paths.items()}
            if not all(paths.values()):
                raise ValueError("Todos os arquivos devem ser selecionados.")

            # Coleta e validação de outros parâmetros
            municipio = self.municipio_combobox.get()
            ano = self.ano_entry.get()
            mes = self.mes_entry.get()
            if not all([municipio, ano, mes]):
                raise ValueError("Município, Ano e Mês devem ser preenchidos.")

            # Define o diretório de saída baseado na localização dos arquivos de entrada
            output_dir = os.path.dirname(paths["licitacoes_html"])

            # Instancia e executa o backend
            backend = PreenchedorQualificacao()
            
            # A função de callback é passada diretamente para o backend
            # A atualização da GUI é feita através do self.after para garantir thread-safety
            def status_update_thread_safe(message, progress):
                self.after(0, self.gui_status_update, message, progress)
            
            backend.execute_all(
                file_paths=paths,
                municipio_nome=municipio,
                exercicio=ano,
                mes=mes,
                status_callback=status_update_thread_safe,
                output_dir=output_dir
            )
            
            # Garante que a mensagem final de sucesso tenha a cor correta
            self.after(0, lambda: self.status_label.configure(text_color="lightgreen"))

        except Exception as e:
            # Em caso de erro, atualiza a GUI com a mensagem
            error_message = f"Erro: {e}"
            self.after(0, self.gui_status_update, error_message, self.progress_bar.get())
            self.after(0, lambda: self.status_label.configure(text_color="orange"))
            traceback.print_exc() # Loga o traceback completo no console para depuração
        finally:
            # Reabilita o botão de início, não importa o que aconteça
            self.run_button.configure(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
