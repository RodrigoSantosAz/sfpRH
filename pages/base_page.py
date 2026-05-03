import customtkinter as ctk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import json
import os

_RECENT_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "recent_files.json")
_MAX_RECENT = 5


def _ler_json():
    if os.path.exists(_RECENT_PATH):
        try:
            with open(_RECENT_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _salvar_json(data):
    with open(_RECENT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class BasePage(ctk.CTkFrame):
    page_key = ""

    def __init__(self, parent, title):
        super().__init__(parent, corner_radius=0)
        self.arquivo_selecionado = ""
        self.df_atual = None
        self._title = title
        self._configurar_estilo_treeview()
        self._build_ui()

    def _configurar_estilo_treeview(self):
        style = ttk.Style()
        style.theme_use("default")
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg = "#2b2b2b" if is_dark else "#ebebeb"
        fg = "#ebebeb" if is_dark else "#1a1a1a"
        style.configure("Preview.Treeview",
            background=bg, foreground=fg, rowheight=24,
            fieldbackground=bg, borderwidth=0, font=("Arial", 10),
        )
        style.configure("Preview.Treeview.Heading",
            background="#1f538d", foreground="white",
            font=("Arial", 10, "bold"), relief="flat",
        )
        style.map("Preview.Treeview",
            background=[("selected", "#1f538d")],
            foreground=[("selected", "white")],
        )

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Row 0 — Header
        header = ctk.CTkFrame(self, height=60, corner_radius=0, fg_color=("gray88", "gray15"))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text=self._title, font=("Arial", 20, "bold")).grid(
            row=0, column=0, padx=25, pady=15, sticky="w"
        )

        # Row 1 — Toolbar
        toolbar = ctk.CTkFrame(self, height=55, fg_color=("gray82", "gray18"))
        toolbar.grid(row=1, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        ctk.CTkButton(
            toolbar, text="📂 Importar", width=130, command=self.importar
        ).pack(side="left", padx=15, pady=10)

        self.label_arquivo = ctk.CTkLabel(
            toolbar, text="Nenhum arquivo selecionado", font=("Arial", 11), text_color="gray60"
        )
        self.label_arquivo.pack(side="left", padx=5)

        self.progress = ctk.CTkProgressBar(toolbar, mode="indeterminate", width=120, height=8)

        ctk.CTkButton(
            toolbar, text="💾 Exportar", width=120, command=self.exportar
        ).pack(side="right", padx=15, pady=10)

        # Row 2 — Barra de arquivos recentes
        self.recent_frame = ctk.CTkFrame(self, height=36, fg_color=("gray78", "gray20"))
        self.recent_frame.grid(row=2, column=0, sticky="ew")
        self.recent_frame.grid_propagate(False)
        self._atualizar_recentes_ui()

        # Row 3 — Conteúdo principal (expande)
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=3, column=0, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        # 3a — Área específica da página (instrução / controles)
        self.content_frame = ctk.CTkFrame(main, fg_color="transparent")
        self.content_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(12, 0))
        self.content_frame.grid_columnconfigure(0, weight=1)
        self._build_content()

        # 3b — Pré-visualização
        preview_container = ctk.CTkFrame(main, fg_color="transparent")
        preview_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        preview_container.grid_columnconfigure(0, weight=1)
        preview_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            preview_container,
            text="Pré-visualização (primeiras 50 linhas)",
            font=("Arial", 11, "bold"), text_color="gray55",
        ).grid(row=0, column=0, sticky="w", pady=(0, 5))

        tree_outer = ctk.CTkFrame(preview_container, fg_color=("gray85", "gray22"))
        tree_outer.grid(row=1, column=0, sticky="nsew")
        tree_outer.grid_columnconfigure(0, weight=1)
        tree_outer.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_outer, style="Preview.Treeview", show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tree_outer, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_outer, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.preview_placeholder = ctk.CTkLabel(
            tree_outer,
            text="Importe uma planilha para visualizar os dados aqui.",
            text_color="gray55", font=("Arial", 12),
        )
        self.preview_placeholder.place(relx=0.5, rely=0.5, anchor="center")

        # Row 4 — Barra de status
        self.status_label = ctk.CTkLabel(
            self, text="", font=("Arial", 11), text_color="gray55"
        )
        self.status_label.grid(row=4, column=0, padx=25, pady=(0, 8), sticky="w")

    def _build_content(self):
        pass

    # ------------------------------------------------------------------ Import

    def importar(self):
        path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if not path:
            return
        nome = path.replace("\\", "/").split("/")[-1]
        self.label_arquivo.configure(text=f"Carregando {nome}…", text_color="gray60")
        self._set_loading(True)
        threading.Thread(target=self._thread_importar, args=(path,), daemon=True).start()

    def _thread_importar(self, path):
        try:
            df = pd.read_excel(path)
            self.after(0, lambda: self._on_importado(path, df))
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self._on_erro_importar(msg))

    def _on_importado(self, path, df):
        self._set_loading(False)
        self.arquivo_selecionado = path
        self.df_atual = df
        nome = path.replace("\\", "/").split("/")[-1]
        self.label_arquivo.configure(text=nome, text_color=("gray10", "gray80"))
        self._validar_e_exibir(df)
        self._salvar_recente(path)
        self._atualizar_recentes_ui()

    def _on_erro_importar(self, msg):
        self._set_loading(False)
        self._set_status(f"Erro ao importar: {msg}", "error")
        messagebox.showerror("Erro ao importar", msg)

    def _validar_e_exibir(self, df):
        faltando = [c for c in self.colunas_necessarias() if c not in df.columns]
        if faltando:
            self._set_status(
                f"Aviso: coluna(s) não encontrada(s): {', '.join(faltando)}", "warning"
            )
        else:
            self._set_status(
                f"Arquivo carregado — {len(df)} linha(s) encontrada(s).", "success"
            )
        self._atualizar_preview(df)

    def _atualizar_preview(self, df):
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, minwidth=80, stretch=False)
        for _, row in df.head(50).iterrows():
            self.tree.insert("", "end", values=[str(v) for v in row])
        self.preview_placeholder.place_forget()

    # ------------------------------------------------------------------ Export

    def exportar(self):
        if self.df_atual is None:
            messagebox.showwarning("Aviso", "Nenhuma planilha importada!")
            return
        destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not destino:
            return
        self._set_loading(True)
        df_copy = self.df_atual.copy()
        threading.Thread(
            target=self._thread_exportar, args=(df_copy, destino), daemon=True
        ).start()

    def _thread_exportar(self, df, destino):
        try:
            df = self.processar(df)
            df.to_excel(destino, index=False)
            self.after(0, self._on_exportado)
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self._on_erro_exportar(msg))

    def _on_exportado(self):
        self._set_loading(False)
        self._set_status("Planilha exportada com sucesso!", "success")

    def _on_erro_exportar(self, msg):
        self._set_loading(False)
        self._set_status(f"Erro ao exportar: {msg}", "error")
        messagebox.showerror("Erro ao exportar", msg)

    # ------------------------------------------------------------------ Overridable

    def processar(self, df):
        return df

    def colunas_necessarias(self):
        return []

    # ------------------------------------------------------------------ Helpers

    def _set_loading(self, ativo):
        if ativo:
            self.progress.pack(side="left", padx=8)
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def _set_status(self, msg, tipo="info"):
        cores = {
            "success": "#2fa843",
            "error":   "#e74c3c",
            "warning": "#e67e22",
            "info":    "gray55",
        }
        self.status_label.configure(text=msg, text_color=cores.get(tipo, "gray55"))

    # ------------------------------------------------------------------ Recentes

    def _salvar_recente(self, path):
        if not self.page_key:
            return
        data = _ler_json()
        lst = data.get(self.page_key, [])
        if path in lst:
            lst.remove(path)
        lst.insert(0, path)
        data[self.page_key] = lst[:_MAX_RECENT]
        _salvar_json(data)

    def _atualizar_recentes_ui(self):
        for w in self.recent_frame.winfo_children():
            w.destroy()

        if not self.page_key:
            return

        recentes = [p for p in _ler_json().get(self.page_key, []) if os.path.exists(p)]

        ctk.CTkLabel(
            self.recent_frame, text="Recentes:",
            font=("Arial", 10, "bold"), text_color="gray55",
        ).pack(side="left", padx=(15, 5), pady=8)

        if not recentes:
            ctk.CTkLabel(
                self.recent_frame, text="—", font=("Arial", 10), text_color="gray50"
            ).pack(side="left", padx=5)
            return

        for path in recentes[:3]:
            nome = path.replace("\\", "/").split("/")[-1]
            ctk.CTkButton(
                self.recent_frame, text=nome,
                font=("Arial", 10), height=22, width=0,
                fg_color="transparent", text_color=("gray20", "gray70"),
                hover_color=("gray72", "gray30"),
                command=lambda p=path: self._carregar_recente(p),
            ).pack(side="left", padx=4, pady=7)

    def _carregar_recente(self, path):
        if not os.path.exists(path):
            messagebox.showwarning("Aviso", f"Arquivo não encontrado:\n{path}")
            self._atualizar_recentes_ui()
            return
        nome = path.replace("\\", "/").split("/")[-1]
        self.label_arquivo.configure(text=f"Carregando {nome}…", text_color="gray60")
        self._set_loading(True)
        threading.Thread(target=self._thread_importar, args=(path,), daemon=True).start()
