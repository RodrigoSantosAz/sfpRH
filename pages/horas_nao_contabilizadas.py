import customtkinter as ctk
from .base_page import BasePage


class HorasNaoContabilizadasPage(BasePage):
    page_key = "horas_nao_contabilizadas"

    def __init__(self, parent):
        super().__init__(parent, "📋  Horas Não-Contabilizadas")

    def _build_content(self):
        ctk.CTkLabel(
            self.content_frame,
            text="Importe uma planilha para identificar horas não-contabilizadas.",
            font=("Arial", 12), text_color="gray55",
        ).grid(row=0, column=0, sticky="w")

    def colunas_necessarias(self):
        return []  # ex: ["Nome", "Horas Registradas", "Horas Sistema"]

    def processar(self, df):
        # Adicione aqui a lógica de identificação de horas não-contabilizadas
        return df
