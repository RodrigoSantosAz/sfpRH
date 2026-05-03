import customtkinter as ctk
from .base_page import BasePage


class FuncionariosAtivosPage(BasePage):
    page_key = "funcionarios_ativos"

    def __init__(self, parent):
        super().__init__(parent, "👥  Funcionários Ativos")

    def _build_content(self):
        ctk.CTkLabel(
            self.content_frame,
            text="Importe uma planilha para filtrar os funcionários ativos.",
            font=("Arial", 12), text_color="gray55",
        ).grid(row=0, column=0, sticky="w")

    def colunas_necessarias(self):
        return []  # ex: ["Nome", "Status", "Data Admissão"]

    def processar(self, df):
        # Adicione aqui a lógica de filtragem de funcionários ativos
        return df
