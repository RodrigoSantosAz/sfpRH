import customtkinter as ctk
from pages.home import HomePage
from pages.horas_extras import HorasExtrasPage
from pages.funcionarios_ativos import FuncionariosAtivosPage
from pages.auxilios_mes import AuxiliosMesPage
from pages.horas_nao_contabilizadas import HorasNaoContabilizadasPage
from pages.documentacao import DocumentacaoPage
from pages.relatorio_ferias import RelatorioFeriasPage

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Planilhas - Departamento Pessoal")
        self.geometry("950x600")
        self.minsize(750, 450)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=210, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)
        self.sidebar.grid_rowconfigure(11, weight=1)

        ctk.CTkLabel(
            self.sidebar, text="SFP RH", font=("Arial", 20, "bold")
        ).grid(row=0, column=0, padx=20, pady=(28, 5))

        ctk.CTkLabel(
            self.sidebar, text="MENU", font=("Arial", 10), text_color="gray55"
        ).grid(row=1, column=0, padx=20, pady=(0, 8))

        nav_items = [
            ("home",                    "🏠   Home"),
            ("horas_extras",            "⏱   Horas Extras"),
            ("funcionarios_ativos",     "👥   Funcionários Ativos"),
            ("auxilios_mes",            "💰   Auxílios do Mês"),
            ("horas_nao_contabilizadas","📋   Horas Não-Cont."),
            ("documentacao",            "📄   Documentação"),
            ("relatorio_ferias",        "🏖   Relatório Férias"),
        ]

        self.buttons = {}
        for i, (key, label) in enumerate(nav_items):
            btn = ctk.CTkButton(
                self.sidebar,
                text=label,
                anchor="w",
                height=38,
                corner_radius=8,
                fg_color="transparent",
                text_color=("gray10", "gray85"),
                hover_color=("gray80", "gray28"),
                command=lambda k=key: self.show_page(k),
            )
            btn.grid(row=i + 2, column=0, padx=10, pady=3, sticky="ew")
            self.buttons[key] = btn

        # --- Área de conteúdo ---
        self.content = ctk.CTkFrame(self, corner_radius=0)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        # --- Páginas ---
        self.pages = {
            "home":                     HomePage(self.content, navigate=self.show_page),
            "horas_extras":             HorasExtrasPage(self.content),
            "funcionarios_ativos":      FuncionariosAtivosPage(self.content),
            "auxilios_mes":             AuxiliosMesPage(self.content),
            "horas_nao_contabilizadas": HorasNaoContabilizadasPage(self.content),
            "documentacao":             DocumentacaoPage(self.content),
            "relatorio_ferias":         RelatorioFeriasPage(self.content),
        }
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

        self.show_page("home")

    def show_page(self, name):
        for page in self.pages.values():
            page.grid_remove()
        page = self.pages[name]
        page.grid()
        if hasattr(page, "on_show"):
            page.after_idle(page.on_show)

        for key, btn in self.buttons.items():
            if key == name:
                btn.configure(fg_color=("gray72", "gray32"), text_color=("black", "white"))
            else:
                btn.configure(fg_color="transparent", text_color=("gray10", "gray85"))


if __name__ == "__main__":
    app = App()
    app.mainloop()
