import tkinter as tk
from tkinter import filedialog
import unicodedata
import re

def ask_file_path(title, filetypes):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return file_path

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    # Converter para minúsculas
    texto = texto.lower()
    # Remover acentos
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')
    # Remover caracteres especiais, mantendo apenas letras, números e espaços
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    # Substituir múltiplos espaços por um único espaço
    texto = re.sub(r'\s+', ' ', texto)
    # Remover espaços no início e no fim
    texto = texto.strip()
    return texto