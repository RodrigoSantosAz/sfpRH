import datetime
import os
import threading
import warnings
from tkinter import filedialog, messagebox

import customtkinter as ctk
import tkinter.ttk as ttk

from .ferias_parser import parse_excel
from .ferias_logica import classificar
from .ferias_pdf import gerar_pdf


class RelatorioFeriasPage(ctk.CTkFrame):
    page_key = "relatorio_ferias"

    def __init__(self, parent):
        super().__init__(parent, corner_radius=0)
        self._pasta = ""
        self._arquivos: list[str] = []
        self._configurar_estilo_treeview()
        self._build_ui()
        self._bind_data_entry()

    # ------------------------------------------------------------------ estilo

    def _configurar_estilo_treeview(self):
        style = ttk.Style()
        style.theme_use("default")
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg = "#2b2b2b" if is_dark else "#ebebeb"
        fg = "#ebebeb" if is_dark else "#1a1a1a"
        style.configure("Ferias.Treeview",
            background=bg, foreground=fg, rowheight=26,
            fieldbackground=bg, borderwidth=0, font=("Arial", 10),
        )
        style.configure("Ferias.Treeview.Heading",
            background="#1f538d", foreground="white",
            font=("Arial", 10, "bold"), relief="flat",
        )
        style.map("Ferias.Treeview",
            background=[("selected", "#1f538d")],
            foreground=[("selected", "white")],
        )

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Row 0 — Header
        header = ctk.CTkFrame(self, height=60, corner_radius=0, fg_color=("gray88", "gray15"))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="🏖  Relatório Férias", font=("Arial", 20, "bold")).grid(
            row=0, column=0, padx=25, pady=15, sticky="w"
        )

        # Row 1 — Painel de configuração
        config = ctk.CTkFrame(self, fg_color=("gray82", "gray18"))
        config.grid(row=1, column=0, sticky="ew")
        config.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(config, text="Pasta dos arquivos:", font=("Arial", 12)).grid(
            row=0, column=0, padx=(20, 10), pady=(18, 8), sticky="w"
        )
        self._entry_pasta = ctk.CTkEntry(
            config,
            placeholder_text="Selecione a pasta com os arquivos .xlsx...",
            font=("Arial", 11),
        )
        self._entry_pasta.grid(row=0, column=1, padx=(0, 10), pady=(18, 8), sticky="ew")
        ctk.CTkButton(
            config, text="📂  Selecionar", width=130, command=self._selecionar_pasta
        ).grid(row=0, column=2, padx=(0, 20), pady=(18, 8))

        ctk.CTkLabel(config, text="Data vencimento:", font=("Arial", 12)).grid(
            row=1, column=0, padx=(20, 10), pady=(0, 16), sticky="w"
        )
        self._entry_data = ctk.CTkEntry(
            config,
            placeholder_text="DD/MM/AAAA",
            width=150,
            font=("Arial", 11),
        )
        self._entry_data.grid(row=1, column=1, padx=(0, 10), pady=(0, 16), sticky="w")

        # Row 2 — Lista de arquivos (expande)
        lista_frame = ctk.CTkFrame(self, fg_color="transparent")
        lista_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(14, 0))
        lista_frame.grid_columnconfigure(0, weight=1)
        lista_frame.grid_rowconfigure(1, weight=1)

        self._label_arquivos = ctk.CTkLabel(
            lista_frame,
            text="Arquivos encontrados: —",
            font=("Arial", 11, "bold"),
            text_color="gray55",
        )
        self._label_arquivos.grid(row=0, column=0, sticky="w", pady=(0, 6))

        outer = ctk.CTkFrame(lista_frame, fg_color=("gray85", "gray22"))
        outer.grid(row=1, column=0, sticky="nsew")
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            outer,
            style="Ferias.Treeview",
            show="headings",
            columns=["arquivo", "registros"],
            selectmode="none",
        )
        self._tree.heading("arquivo", text="Arquivo")
        self._tree.heading("registros", text="Registros")
        self._tree.column("arquivo", stretch=True, minwidth=200)
        self._tree.column("registros", width=110, stretch=False, anchor="center")
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._placeholder = ctk.CTkLabel(
            outer,
            text="Selecione uma pasta para listar os arquivos .xlsx disponíveis.",
            text_color="gray55",
            font=("Arial", 12),
        )
        self._placeholder.place(relx=0.5, rely=0.5, anchor="center")

        # Row 3 — Barra inferior
        bottom = ctk.CTkFrame(self, height=55, fg_color=("gray82", "gray18"))
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.grid_propagate(False)

        self._btn_executar = ctk.CTkButton(
            bottom,
            text="▶  Gerar Relatórios",
            width=170,
            state="disabled",
            command=self._executar,
        )
        self._btn_executar.pack(side="right", padx=20, pady=10)

        self._status = ctk.CTkLabel(bottom, text="", font=("Arial", 11), text_color="gray55")
        self._status.pack(side="left", padx=20, pady=10)

    # ------------------------------------------------------------------ lógica

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta com arquivos .xlsx")
        if not pasta:
            return
        self._pasta = pasta
        self._entry_pasta.delete(0, "end")
        self._entry_pasta.insert(0, pasta)
        self._atualizar_lista()

    def _atualizar_lista(self):
        self._tree.delete(*self._tree.get_children())
        arquivos = sorted([
            f for f in os.listdir(self._pasta)
            if f.lower().endswith(".xlsx") and not f.startswith("~$")
        ])
        self._arquivos = arquivos

        if not arquivos:
            self._label_arquivos.configure(text="Arquivos encontrados: nenhum .xlsx nesta pasta")
            self._btn_executar.configure(state="disabled")
            self._placeholder.configure(text="Nenhum arquivo .xlsx encontrado nesta pasta.")
            self._placeholder.place(relx=0.5, rely=0.5, anchor="center")
            self._status.configure(text="")
            return

        self._placeholder.place_forget()
        self._label_arquivos.configure(text=f"Arquivos encontrados: {len(arquivos)} .xlsx")
        self._btn_executar.configure(state="normal")
        self._status.configure(text="")

        # Preenche a lista; contagem de registros é carregada em thread
        for arq in arquivos:
            self._tree.insert("", "end", iid=arq, values=(arq, "lendo..."))

        threading.Thread(target=self._contar_registros, args=(list(arquivos),), daemon=True).start()

    def _contar_registros(self, arquivos: list[str]):
        for arq in arquivos:
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    registros = parse_excel(os.path.join(self._pasta, arq))
                total = len([r for r in registros if r.dias_pendentes > 0])
                self.after(0, lambda a=arq, t=total: self._tree.set(a, "registros", str(t)))
            except Exception:
                self.after(0, lambda a=arq: self._tree.set(a, "registros", "erro"))

    # ------------------------------------------------------------------ campo data

    def _bind_data_entry(self):
        tk_e = self._entry_data._entry
        tk_e.bind("<KeyPress>", self._on_data_keypress)

    def _on_data_keypress(self, event):
        # Permite navegação e atalhos de teclado
        if event.keysym in ("Left", "Right", "Home", "End", "Tab", "Return"):
            return
        if event.state & 0x4:  # Ctrl (copiar/colar/selecionar)
            return
        if event.keysym == "BackSpace":
            self._data_backspace()
            return "break"
        if event.keysym == "Delete":
            self._data_delete()
            return "break"
        if not event.char or not event.char.isdigit():
            return "break"
        self._data_insert(event.char)
        return "break"

    @staticmethod
    def _data_format(digitos: str) -> str:
        d = digitos[:8]
        if len(d) > 4:
            return d[:2] + "/" + d[2:4] + "/" + d[4:]
        if len(d) > 2:
            return d[:2] + "/" + d[2:]
        return d

    def _data_apply(self, digitos: str, cursor_apos_n_digitos: int):
        """Aplica o texto formatado e posiciona o cursor após o n-ésimo dígito."""
        formatado = self._data_format(digitos)
        tk_e = self._entry_data._entry
        tk_e.delete(0, "end")
        tk_e.insert(0, formatado)
        # Calcula posição do cursor: após o n-ésimo dígito no texto formatado
        pos = len(formatado)
        count = 0
        for i, c in enumerate(formatado):
            if c.isdigit():
                count += 1
                if count == cursor_apos_n_digitos:
                    pos = i + 1
                    break
        # Pula a barra se o cursor cair sobre ela
        if pos < len(formatado) and formatado[pos] == "/":
            pos += 1
        tk_e.icursor(pos)

    def _data_insert(self, char: str):
        tk_e = self._entry_data._entry
        texto = tk_e.get()
        pos = tk_e.index("insert")
        # Remove seleção se houver
        sel_antes = None
        try:
            s1 = tk_e.index("sel.first")
            s2 = tk_e.index("sel.last")
            todos = [c for c in texto if c.isdigit()]
            antes = sum(1 for c in texto[:s1] if c.isdigit())
            em = sum(1 for c in texto[s1:s2] if c.isdigit())
            del todos[antes:antes + em]
            sel_antes = antes
            texto = "".join(todos)
            pos_digito = antes
        except Exception:
            todos = [c for c in texto if c.isdigit()]
            pos_digito = sum(1 for c in texto[:pos] if c.isdigit())

        if sel_antes is None:
            todos = [c for c in texto if c.isdigit()]

        if len(todos) >= 8:
            return
        todos.insert(pos_digito, char)
        self._data_apply("".join(todos), pos_digito + 1)

    def _data_backspace(self):
        tk_e = self._entry_data._entry
        texto = tk_e.get()
        try:
            s1 = tk_e.index("sel.first")
            s2 = tk_e.index("sel.last")
            todos = [c for c in texto if c.isdigit()]
            antes = sum(1 for c in texto[:s1] if c.isdigit())
            em = sum(1 for c in texto[s1:s2] if c.isdigit())
            del todos[antes:antes + em]
            self._data_apply("".join(todos), antes)
            return
        except Exception:
            pass
        pos = tk_e.index("insert")
        todos = [c for c in texto if c.isdigit()]
        digitos_antes = sum(1 for c in texto[:pos] if c.isdigit())
        if digitos_antes == 0:
            return
        del todos[digitos_antes - 1]
        self._data_apply("".join(todos), digitos_antes - 1)

    def _data_delete(self):
        tk_e = self._entry_data._entry
        texto = tk_e.get()
        try:
            s1 = tk_e.index("sel.first")
            s2 = tk_e.index("sel.last")
            todos = [c for c in texto if c.isdigit()]
            antes = sum(1 for c in texto[:s1] if c.isdigit())
            em = sum(1 for c in texto[s1:s2] if c.isdigit())
            del todos[antes:antes + em]
            self._data_apply("".join(todos), antes)
            return
        except Exception:
            pass
        pos = tk_e.index("insert")
        todos = [c for c in texto if c.isdigit()]
        digitos_antes = sum(1 for c in texto[:pos] if c.isdigit())
        if digitos_antes >= len(todos):
            return
        del todos[digitos_antes]
        self._data_apply("".join(todos), digitos_antes)

    def _get_data_vencimento(self) -> datetime.date | None:
        texto = self._entry_data.get().strip()
        try:
            return datetime.datetime.strptime(texto, "%d/%m/%Y").date()
        except ValueError:
            return None

    def _executar(self):
        if not self._pasta or not self._arquivos:
            messagebox.showwarning("Aviso", "Selecione uma pasta com arquivos .xlsx.")
            return
        data_venc = self._get_data_vencimento()
        if data_venc is None:
            messagebox.showwarning(
                "Data inválida",
                "Informe a data de vencimento no formato DD/MM/AAAA.\nExemplo: 30/06/2026",
            )
            return
        ProgressDialog(self, self._pasta, list(self._arquivos), data_venc)


# ====================================================================== popup

class ProgressDialog(ctk.CTkToplevel):

    def __init__(self, parent, pasta: str, arquivos: list[str], data_vencimento: datetime.date):
        super().__init__(parent)
        self.title("Gerando relatórios...")
        self.geometry("600x420")
        self.minsize(500, 320)
        self.grab_set()

        self._pasta = pasta
        self._arquivos = arquivos
        self._data_venc = data_vencimento
        self._widgets: dict[str, tuple] = {}  # arquivo → (bar, status_label)

        self._build_ui()
        self.after(150, self._iniciar)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self, text="Processando arquivos...", font=("Arial", 14, "bold")
        ).grid(row=0, column=0, padx=20, pady=(18, 8), sticky="w")

        scroll = ctk.CTkScrollableFrame(self, fg_color=("gray90", "gray20"))
        scroll.grid(row=1, column=0, padx=20, pady=5, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        for i, arq in enumerate(self._arquivos):
            bloco = ctk.CTkFrame(scroll, fg_color="transparent")
            bloco.grid(row=i, column=0, sticky="ew", pady=8)
            bloco.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                bloco, text=f"📄  {arq}", font=("Arial", 11), anchor="w"
            ).grid(row=0, column=0, columnspan=2, sticky="w", padx=4)

            barra = ctk.CTkProgressBar(bloco, height=10)
            barra.set(0)
            barra.grid(row=1, column=0, padx=4, pady=(5, 0), sticky="ew")

            status = ctk.CTkLabel(
                bloco,
                text="Aguardando...",
                font=("Arial", 10),
                text_color="gray55",
                width=110,
                anchor="w",
            )
            status.grid(row=1, column=1, padx=(10, 4))

            self._widgets[arq] = (barra, status)

        self._btn_fechar = ctk.CTkButton(
            self, text="Fechar", state="disabled", width=110, command=self.destroy
        )
        self._btn_fechar.grid(row=2, column=0, padx=20, pady=15, sticky="e")

    # ------------------------------------------------------------------ execução

    def _iniciar(self):
        threading.Thread(target=self._processar_todos, daemon=True).start()

    def _processar_todos(self):
        hoje = datetime.date.today()
        nome_saida = f"{hoje.strftime('%Y-%m-%d')}_relatorios_vencidos"
        pasta_saida = os.path.join(self._pasta, nome_saida)
        os.makedirs(pasta_saida, exist_ok=True)

        for arq in self._arquivos:
            self._set_status(arq, "Lendo arquivo...", "gray55", 0.15)
            try:
                caminho = os.path.join(self._pasta, arq)
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    registros = parse_excel(caminho)

                self._set_status(arq, "Classificando...", "gray55", 0.45)
                pessoas = classificar(registros, self._data_venc, hoje)

                self._set_status(arq, "Gerando PDF...", "gray55", 0.75)
                nome_pdf = os.path.splitext(arq)[0] + "_VENCIDOS.pdf"
                destino = os.path.join(pasta_saida, nome_pdf)
                gerar_pdf(pessoas, self._data_venc, destino)

                self._set_status(arq, f"✓  Concluído ({len(pessoas)} pessoa(s))", "#2fa843", 1.0)

            except ImportError:
                self._set_status(
                    arq,
                    "✗  fpdf2 não instalado",
                    "#e74c3c",
                    0.0,
                )
            except Exception as e:
                msg = str(e)[:40]
                self._set_status(arq, f"✗  {msg}", "#e74c3c", 0.0)

        self.after(0, lambda: self.title("Relatórios gerados!"))
        self.after(0, lambda: self._btn_fechar.configure(state="normal"))

    def _set_status(self, arq: str, texto: str, cor: str, progresso: float):
        barra, label = self._widgets[arq]
        self.after(0, lambda b=barra, p=progresso: b.set(p))
        self.after(0, lambda lb=label, t=texto, c=cor: lb.configure(text=t, text_color=c))
