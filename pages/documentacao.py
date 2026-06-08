import os
import re
import json
import threading
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document

_TEMPLATE_PROFESSOR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "arquivos", "Relatorio Professor.docx"
)
_HISTORICO_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "arquivos", "sugestoes_professor.json"
)


# ---------------------------------------------------------------------------
# Histórico de sugestões (CURSO / INSTITUICAO)
# ---------------------------------------------------------------------------
def _carregar_historico() -> dict:
    if os.path.exists(_HISTORICO_PATH):
        try:
            with open(_HISTORICO_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _salvar_historico(data: dict):
    with open(_HISTORICO_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Mapeamentos e definição dos campos
# ---------------------------------------------------------------------------
_GENERO_MAP = {
    "Masculino": "o servidor",
    "Feminino":  "a servidora",
}

_PORCENTAGEM_MAP  = {"1": "15%",          "2": "25%",      "3": "40%"}
_NIVEL_CURSO_MAP  = {"1": "Especialização","2": "Mestrado",  "3": "Doutorado"}

_MESES_PT = [
    "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro",
]
_DIAS = [str(d) for d in range(1, 32)]
_ANOS = [str(a) for a in range(1990, 2036)]

_hoje = datetime.now()
_DIA_ATUAL = str(_hoje.day)
_MES_ATUAL = _MESES_PT[_hoje.month - 1]
_ANO_ATUAL = str(_hoje.year)

# Formato: (tipo, placeholder, label, extra)
#   tipo "entry"        → campo de texto livre
#   tipo "dropdown"     → CTkOptionMenu; extra = lista de opções
#   tipo "dropdown_map" → CTkOptionMenu com mapeamento; extra = dict {exibição: valor}
#   tipo "auto"         → calculado automaticamente, sem input do usuário
_CAMPOS_PROFESSOR = [
    ("Documento", [
        ("entry",        "{{NUMERO_DOC}}", "Número do Documento",     None),
    ]),
    ("Dados do Servidor", [
        ("entry",        "{{NOME}}",        "Nome",                    None),
        ("dropdown_map", "{{GENERO}}",      "Gênero",                  _GENERO_MAP),
        ("dropdown",     "{{NIVEL}}",       "Nível Atual",             ["1", "2", "3"]),
        ("date",         "{{DATA_INICIO}}", "Data de Início no Cargo", None),
    ]),
    ("Alteração de Nível", [
        ("auto",         "{{NIVEL+1}}",     "Próximo Nível",           None),
        ("auto",         "{{PORCENTAGEM}}", "Porcentagem do Acréscimo",None),
        ("auto",         "{{NIVEL_CURSO}}", "Tipo de Formação",        None),
        ("combobox",     "{{CURSO}}",       "Nome do Curso",           None),
        ("combobox",     "{{INSTITUICAO}}", "Instituição",             None),
    ]),
    ("Data do Parecer", [
        ("dropdown", "{{DATA_DIA}}", "Dia", (_DIAS,     _DIA_ATUAL)),
        ("dropdown", "{{DATA_MES}}", "Mês", (_MESES_PT, _MES_ATUAL)),
        ("dropdown", "{{DATA_ANO}}", "Ano", (_ANOS,     _ANO_ATUAL)),
    ]),
]

# Campos obrigatórios por tipo
_ENTRY_PLACEHOLDERS = {
    ph for _, campos in _CAMPOS_PROFESSOR
    for tipo, ph, *_ in campos
    if tipo in ("entry", "combobox")
}
_DATE_PLACEHOLDERS = {
    ph for _, campos in _CAMPOS_PROFESSOR
    for tipo, ph, *_ in campos
    if tipo == "date"
}


# ---------------------------------------------------------------------------
# Widget: entrada com máscara DD/MM/AAAA
# ---------------------------------------------------------------------------
class _DateEntry(ctk.CTkEntry):
    """CTkEntry com máscara automática DD/MM/AAAA e validação de data real."""

    def __init__(self, master, **kwargs):
        self._var = ctk.StringVar()
        super().__init__(master, textvariable=self._var, placeholder_text="DD/MM/AAAA", **kwargs)
        self._updating = False
        self._var.trace_add("write", self._formatar)

    def _formatar(self, *_):
        if self._updating:
            return
        self._updating = True
        raw = self._var.get()
        digits = re.sub(r"\D", "", raw)[:8]
        formatted = digits[:2]
        if len(digits) >= 3:
            formatted += "/" + digits[2:4]
        if len(digits) >= 5:
            formatted += "/" + digits[4:8]
        if formatted != raw:
            self._var.set(formatted)
            self.after(0, lambda: self.icursor("end"))
        self._updating = False

    def get(self) -> str:
        return self._var.get()

    def is_valid(self) -> bool:
        try:
            datetime.strptime(self._var.get(), "%d/%m/%Y")
            return True
        except ValueError:
            return False


# ---------------------------------------------------------------------------
# Página principal de Documentação
# ---------------------------------------------------------------------------
class DocumentacaoPage(ctk.CTkFrame):
    page_key = "documentacao"

    _TIPOS = ["Professor", "Outros Cargos"]

    def __init__(self, parent):
        super().__init__(parent, corner_radius=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _build_ui(self):
        header = ctk.CTkFrame(self, height=60, corner_radius=0, fg_color=("gray88", "gray15"))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            header, text="Documentação", font=("Arial", 20, "bold")
        ).grid(row=0, column=0, padx=25, pady=15, sticky="w")

        ctk.CTkLabel(
            header, text="Tipo de documento:", font=("Arial", 12), text_color="gray55"
        ).grid(row=0, column=2, padx=(0, 8), pady=15, sticky="e")

        self._dropdown = ctk.CTkOptionMenu(
            header,
            values=self._TIPOS,
            width=180,
            command=self._on_tipo_alterado,
        )
        self._dropdown.grid(row=0, column=3, padx=(0, 20), pady=15, sticky="e")

        self._content_area = ctk.CTkFrame(self, fg_color="transparent")
        self._content_area.grid(row=1, column=0, sticky="nsew")
        self._content_area.grid_columnconfigure(0, weight=1)
        self._content_area.grid_rowconfigure(0, weight=1)

        self._sub_paginas = {
            "Professor":     _SubPaginaProfessor(self._content_area),
            "Outros Cargos": _SubPaginaOutrosCargos(self._content_area),
        }
        for sub in self._sub_paginas.values():
            sub.grid(row=0, column=0, sticky="nsew")

        self._mostrar("Professor")

    def _on_tipo_alterado(self, valor):
        self._mostrar(valor)

    def _mostrar(self, tipo):
        for sub in self._sub_paginas.values():
            sub.grid_remove()
        self._sub_paginas[tipo].grid()


# ---------------------------------------------------------------------------
# Sub-página: Professor
# ---------------------------------------------------------------------------
class _SubPaginaProfessor(ctk.CTkFrame):

    def __init__(self, parent):
        super().__init__(parent, corner_radius=0, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._getters: dict[str, callable] = {}
        self._date_widgets: dict[str, _DateEntry] = {}
        self._combos: dict[str, ctk.CTkComboBox] = {}
        self._nivel_auto_label: ctk.CTkLabel | None = None
        self._porcentagem_auto_label: ctk.CTkLabel | None = None
        self._nivel_curso_auto_label: ctk.CTkLabel | None = None

        self._build_toolbar()
        self._build_form()
        self._build_statusbar()

    # ------------------------------------------------------------------ UI

    def _build_toolbar(self):
        bar = ctk.CTkFrame(self, height=55, fg_color=("gray82", "gray18"))
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_propagate(False)

        self._progress = ctk.CTkProgressBar(bar, mode="indeterminate", width=120, height=8)

        ctk.CTkButton(
            bar, text="💾 Exportar .docx", width=150, command=self._exportar
        ).pack(side="right", padx=15, pady=10)

    def _build_form(self):
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.grid(row=1, column=0, sticky="nsew", padx=24, pady=(16, 0))
        scroll.grid_columnconfigure(1, weight=1)

        row = 0
        for grupo, campos in _CAMPOS_PROFESSOR:
            ctk.CTkLabel(
                scroll,
                text=grupo.upper(),
                font=("Arial", 10, "bold"),
                text_color="gray50",
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(16, 4))
            row += 1

            for tipo, ph, label, extra in campos:
                ctk.CTkLabel(
                    scroll, text=label, font=("Arial", 12), anchor="w"
                ).grid(row=row, column=0, sticky="w", padx=(8, 16), pady=4)

                if tipo == "entry":
                    widget = ctk.CTkEntry(scroll, height=32, font=("Arial", 12))
                    widget.grid(row=row, column=1, sticky="ew", pady=4)
                    self._getters[ph] = widget.get

                elif tipo == "date":
                    widget = _DateEntry(scroll, height=32, font=("Arial", 12))
                    widget.grid(row=row, column=1, sticky="ew", pady=4)
                    self._getters[ph] = widget.get
                    self._date_widgets[ph] = widget

                elif tipo == "combobox":
                    historico = _carregar_historico()
                    sugestoes = historico.get(ph, [])
                    widget = ctk.CTkComboBox(
                        scroll, values=sugestoes, height=32, font=("Arial", 12)
                    )
                    widget.set("")
                    widget.grid(row=row, column=1, sticky="ew", pady=4)
                    self._getters[ph] = widget.get
                    self._combos[ph] = widget

                elif tipo == "dropdown":
                    opcoes, default = extra if isinstance(extra, tuple) else (extra, extra[0])
                    widget = ctk.CTkOptionMenu(
                        scroll, values=opcoes, height=32, font=("Arial", 12)
                    )
                    widget.set(default)
                    widget.grid(row=row, column=1, sticky="ew", pady=4)
                    self._getters[ph] = widget.get
                    if ph == "{{NIVEL}}":
                        widget.configure(command=self._on_nivel_alterado)

                elif tipo == "dropdown_map":
                    opcoes = list(extra.keys())
                    widget = ctk.CTkOptionMenu(
                        scroll, values=opcoes, height=32, font=("Arial", 12)
                    )
                    widget.set(opcoes[0])
                    widget.grid(row=row, column=1, sticky="ew", pady=4)
                    mapa = extra
                    self._getters[ph] = lambda w=widget, m=mapa: m[w.get()]

                elif tipo == "auto":
                    _iniciais = {
                        "{{NIVEL+1}}":     "2",
                        "{{PORCENTAGEM}}": _PORCENTAGEM_MAP["1"],
                        "{{NIVEL_CURSO}}": _NIVEL_CURSO_MAP["1"],
                    }
                    lbl = ctk.CTkLabel(
                        scroll,
                        text=_iniciais.get(ph, ""),
                        font=("Arial", 12),
                        text_color="gray55",
                        anchor="w",
                    )
                    lbl.grid(row=row, column=1, sticky="w", padx=6, pady=4)

                    if ph == "{{NIVEL+1}}":
                        self._nivel_auto_label = lbl
                    elif ph == "{{PORCENTAGEM}}":
                        self._porcentagem_auto_label = lbl
                    elif ph == "{{NIVEL_CURSO}}":
                        self._nivel_curso_auto_label = lbl

                    self._getters[ph] = lambda w=lbl: w.cget("text")

                row += 1

    def _build_statusbar(self):
        self._status = ctk.CTkLabel(
            self, text="", font=("Arial", 11), text_color="gray55"
        )
        self._status.grid(row=2, column=0, padx=25, pady=(4, 8), sticky="w")

    # ------------------------------------------------------------------ Callbacks

    def _on_nivel_alterado(self, valor: str):
        if self._nivel_auto_label:
            self._nivel_auto_label.configure(text=str(int(valor) + 1))
        if self._porcentagem_auto_label:
            self._porcentagem_auto_label.configure(text=_PORCENTAGEM_MAP[valor])
        if self._nivel_curso_auto_label:
            self._nivel_curso_auto_label.configure(text=_NIVEL_CURSO_MAP[valor])

    # ------------------------------------------------------------------ Export

    def _exportar(self):
        if not os.path.exists(_TEMPLATE_PROFESSOR):
            messagebox.showerror("Erro", f"Template não encontrado:\n{_TEMPLATE_PROFESSOR}")
            return

        substituicoes = {ph: getter() for ph, getter in self._getters.items()}

        vazios = [ph for ph in _ENTRY_PLACEHOLDERS if not substituicoes.get(ph, "").strip()]
        datas_invalidas = [ph for ph, w in self._date_widgets.items() if not w.is_valid()]
        problemas = vazios + datas_invalidas
        if problemas:
            labels_problema = [
                label + (" (data inválida)" if ph in datas_invalidas else "")
                for _, campos in _CAMPOS_PROFESSOR
                for tipo, ph, label, *_ in campos
                if ph in problemas
            ]
            messagebox.showwarning(
                "Campos inválidos",
                "Corrija os campos antes de exportar:\n• " + "\n• ".join(labels_problema),
            )
            return

        destino = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")],
            initialfile="Relatorio Professor_editado.docx",
        )
        if not destino:
            return

        self._set_loading(True)
        threading.Thread(
            target=self._thread_exportar,
            args=(_TEMPLATE_PROFESSOR, substituicoes, destino),
            daemon=True,
        ).start()

    def _thread_exportar(self, origem, subs, destino):
        try:
            doc = Document(origem)
            for para in doc.paragraphs:
                _substituir_paragrafo(para, subs)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            _substituir_paragrafo(para, subs)
            doc.save(destino)
            self.after(0, lambda: self._on_exportado(destino))
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self._on_erro(msg))

    def _on_exportado(self, destino):
        self._set_loading(False)
        self._set_status(f"Exportado com sucesso: {os.path.basename(destino)}", "success")
        self.after(200, self._verificar_novas_sugestoes)

    def _verificar_novas_sugestoes(self):
        historico = _carregar_historico()
        novos = {}
        for ph, combo in self._combos.items():
            valor = combo.get().strip()
            if valor and valor not in historico.get(ph, []):
                label = next(
                    label for _, campos in _CAMPOS_PROFESSOR
                    for tipo, p, label, *_ in campos if p == ph
                )
                novos[ph] = (label, valor)

        if not novos:
            return

        linhas = "\n".join(f"  • {label}: {valor!r}" for _, (label, valor) in novos.items())
        if messagebox.askyesno(
            "Salvar sugestões",
            f"Deseja salvar para sugestões futuras?\n\n{linhas}",
        ):
            for ph, (_, valor) in novos.items():
                historico.setdefault(ph, []).insert(0, valor)
            _salvar_historico(historico)
            for ph, combo in self._combos.items():
                combo.configure(values=historico.get(ph, []))

    def _on_erro(self, msg):
        self._set_loading(False)
        self._set_status(f"Erro: {msg}", "error")
        messagebox.showerror("Erro ao exportar", msg)

    # ------------------------------------------------------------------ Helpers

    def _set_loading(self, ativo):
        if ativo:
            self._progress.pack(side="left", padx=8)
            self._progress.start()
        else:
            self._progress.stop()
            self._progress.pack_forget()

    def _set_status(self, msg, tipo="info"):
        cores = {"success": "#2fa843", "error": "#e74c3c", "warning": "#e67e22", "info": "gray55"}
        self._status.configure(text=msg, text_color=cores.get(tipo, "gray55"))


# ---------------------------------------------------------------------------
# Sub-página: Outros Cargos (em desenvolvimento)
# ---------------------------------------------------------------------------
class _SubPaginaOutrosCargos(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=0, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text="📄  Funcionalidade Outros Cargos — em desenvolvimento",
            font=("Arial", 14),
            text_color="gray50",
        ).grid(row=0, column=0)


# ---------------------------------------------------------------------------
# Utilitário: substituição preservando formatação dos runs
# ---------------------------------------------------------------------------
def _substituir_paragrafo(para, subs: dict):
    texto = "".join(r.text for r in para.runs)
    novo = texto
    for chave, valor in subs.items():
        novo = novo.replace(chave, valor)
    if novo != texto and para.runs:
        para.runs[0].text = novo
        for r in para.runs[1:]:
            r.text = ""
