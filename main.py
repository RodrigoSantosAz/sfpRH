import customtkinter as ctk
from tkinter import filedialog
import pandas as pd

# Configuração visual do app
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Janela principal
app = ctk.CTk()
app.title("Gerador de Planilhas - Departamento Pessoal")
app.geometry("800x500")

# Variável que vai guardar o caminho da planilha importada
arquivo_selecionado = ""

# Função para importar a planilha
def importar_planilha():
    global arquivo_selecionado
    arquivo_selecionado = filedialog.askopenfilename(
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo_selecionado:
        label_arquivo.configure(text=f"Arquivo: {arquivo_selecionado.split('/')[-1]}")
        print(f"Planilha carregada: {arquivo_selecionado}")

# Função para processar e exportar
def exportar_planilha():
    if not arquivo_selecionado:
        print("Nenhuma planilha importada!")
        return

    df = pd.read_excel(arquivo_selecionado)

    # ✏️ Aqui você vai adicionar suas transformações futuramente

    destino = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")]
    )
    if destino:
        df.to_excel(destino, index=False)
        print("Planilha exportada com sucesso!")

# Interface — Título
titulo = ctk.CTkLabel(app, text="Gerador de Planilhas", font=("Arial", 22, "bold"))
titulo.pack(pady=20)

# Botão importar
btn_importar = ctk.CTkButton(app, text="📂 Importar Planilha", command=importar_planilha)
btn_importar.pack(pady=10)

# Label que mostra o arquivo carregado
label_arquivo = ctk.CTkLabel(app, text="Nenhum arquivo selecionado", font=("Arial", 12))
label_arquivo.pack(pady=5)

# Botão exportar
btn_exportar = ctk.CTkButton(app, text="💾 Exportar Planilha", command=exportar_planilha)
btn_exportar.pack(pady=10)

# Inicia o programa
app.mainloop()