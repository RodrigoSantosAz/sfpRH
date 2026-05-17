import customtkinter as ctk


class HomePage(ctk.CTkFrame):
    def __init__(self, parent, navigate=None):
        super().__init__(parent, corner_radius=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        header = ctk.CTkFrame(self, height=60, corner_radius=0, fg_color=("gray88", "gray15"))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="🏠  Home", font=("Arial", 20, "bold")).grid(
            row=0, column=0, padx=25, pady=15, sticky="w"
        )

        # Conteúdo rolável
        content = ctk.CTkScrollableFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=30, pady=20)
        content.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            content, text="Bem-vindo ao Sistema de RH",
            font=("Arial", 22, "bold"),
        ).grid(row=0, column=0, columnspan=2, pady=(10, 4), sticky="w")

        ctk.CTkLabel(
            content,
            text="Selecione uma funcionalidade no menu lateral para começar.",
            font=("Arial", 13), text_color="gray55",
        ).grid(row=1, column=0, columnspan=2, pady=(0, 24), sticky="w")

        features = [
            (
                "⏱  Horas Extras",
                "Processa e calcula as horas extras dos funcionários a partir de uma planilha de ponto.",
                "horas_extras",
            ),
            (
                "👥  Funcionários Ativos",
                "Filtra e lista os funcionários com status ativo no período informado.",
                "funcionarios_ativos",
            ),
            (
                "💰  Auxílios do Mês",
                "Consolida os auxílios (alimentação, transporte, etc.) do mês por funcionário.",
                "auxilios_mes",
            ),
            (
                "📋  Horas Não-Cont.",
                "Identifica horas de trabalho registradas que não foram contabilizadas no sistema.",
                "horas_nao_contabilizadas",
            ),
        ]

        _cor_normal = ("gray83", "gray22")
        _cor_hover  = ("gray75", "gray30")

        for i, (titulo, desc, page_key) in enumerate(features):
            row = (i // 2) + 2
            col = i % 2
            card = ctk.CTkFrame(content, fg_color=_cor_normal, corner_radius=10)
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            card.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                card, text=titulo, font=("Arial", 14, "bold")
            ).grid(row=0, column=0, padx=15, pady=(15, 4), sticky="w")

            ctk.CTkLabel(
                card, text=desc, font=("Arial", 11), text_color="gray55",
                wraplength=260, justify="left",
            ).grid(row=1, column=0, padx=15, pady=(0, 15), sticky="w")

            if navigate:
                def _bind(c, k):
                    def _click(_):
                        navigate(k)

                    def _nter(_):
                        c.configure(fg_color=_cor_hover)

                    def _leave(_):
                        c.configure(fg_color=_cor_normal)

                    for w in [c] + list(c.winfo_children()):
                        w.bind("<Button-1>", _click)
                        w.bind("<Enter>", _nter)
                        w.bind("<Leave>", _leave)
                        try:
                            w.configure(cursor="hand2")
                        except Exception:
                            pass

                _bind(card, page_key)
