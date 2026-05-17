from .wip_page import WIPPage


class HorasNaoContabilizadasPage(WIPPage):
    page_key = "horas_nao_contabilizadas"

    def __init__(self, parent):
        super().__init__(parent)
