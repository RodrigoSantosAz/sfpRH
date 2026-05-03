import customtkinter as ctk
from .base_page import BasePage


class HorasExtrasPage(BasePage):
    page_key = "horas_extras"

    def __init__(self, parent):
        super().__init__(parent, "⏱  Horas Extras")

    def _build_content(self):
        ctk.CTkLabel(
            self.content_frame,
            text="Importe uma planilha de ponto para processar as horas extras.",
            font=("Arial", 12), text_color="gray55",
        ).grid(row=0, column=0, sticky="w")

    def colunas_necessarias(self):
        return []  # ex: ["Nome", "Data", "Entrada", "Saída"]

    def processar(self, df):
        # Adicione aqui a lógica de cálculo de horas extras
        return df
