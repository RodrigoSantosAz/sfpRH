import customtkinter as ctk
from .base_page import BasePage


class AuxiliosMesPage(BasePage):
    page_key = "auxilios_mes"

    def __init__(self, parent):
        super().__init__(parent, "💰  Auxílios do Mês")

    def _build_content(self):
        ctk.CTkLabel(
            self.content_frame,
            text="Importe uma planilha para processar os auxílios do mês.",
            font=("Arial", 12), text_color="gray55",
        ).grid(row=0, column=0, sticky="w")

    def colunas_necessarias(self):
        return []  # ex: ["Nome", "Alimentação", "Transporte"]

    def processar(self, df):
        # Adicione aqui a lógica de cálculo de auxílios
        return df
