import json
import os
import re
import threading
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from .base_page import BasePage

COLUNAS = ["CPF", "MATRÍCULA", "NOME", "EVENTO", "QTD", "VALOR", "MÊS ANTERIOR", "ÓRGAO"]

MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_HISTORICO_PATH = os.path.join(_ROOT, "historico_horas_extras.json")
_PREFS_PATH = os.path.join(_ROOT, "prefs.json")


def _ler_prefs():
    if os.path.exists(_PREFS_PATH):
        try:
            with open(_PREFS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _salvar_pref(chave, valor):
    prefs = _ler_prefs()
    prefs[chave] = valor
    try:
        with open(_PREFS_PATH, "w", encoding="utf-8") as f:
            json.dump(prefs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class HorasExtrasPage(BasePage):
    page_key = "horas_extras"

    def __init__(self, parent):
        self.mes_atual = 0
        self.ano_atual = 0
        super().__init__(parent, "⏱  Horas Extras")

    def _build_content(self):
        ctk.CTkLabel(
            self.content_frame,
            text="Importe uma planilha de horas extras para consolidar e calcular as diferenças por colaborador.",
            font=("Arial", 12), text_color="gray55",
        ).grid(row=0, column=0, sticky="w")

    def _celula_formato(self, col, val):
        if col == "DIFERENÇA MES ANTERIOR":
            try:
                return f"{float(val):.2f}%".replace(".", ",")
            except (TypeError, ValueError):
                return str(val)
        if col in ("TOTAL VALOR MES", "MES ANTERIOR"):
            try:
                return "R$ " + f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except (TypeError, ValueError):
                return str(val)
        return str(val)

    def colunas_necessarias(self):
        return COLUNAS

    # ------------------------------------------------------------------ Importar

    def importar(self):
        self._abrir_dialog_importar()

    def _abrir_dialog_importar(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Importar Planilha de Horas Extras")
        dialog.geometry("460x260")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.grid_columnconfigure(1, weight=1)

        now = datetime.now()

        ctk.CTkLabel(dialog, text="Mês:", anchor="w").grid(
            row=0, column=0, padx=20, pady=(20, 8), sticky="w"
        )
        mes_var = ctk.StringVar(value=MESES[now.month - 1])
        ctk.CTkOptionMenu(dialog, values=MESES, variable=mes_var, width=180).grid(
            row=0, column=1, padx=(0, 20), pady=(20, 8), sticky="w", columnspan=2
        )

        ctk.CTkLabel(dialog, text="Ano:", anchor="w").grid(
            row=1, column=0, padx=20, pady=8, sticky="w"
        )
        ano_var = ctk.StringVar(value=str(now.year))
        ctk.CTkEntry(dialog, textvariable=ano_var, width=100).grid(
            row=1, column=1, padx=(0, 20), pady=8, sticky="w", columnspan=2
        )

        ctk.CTkLabel(dialog, text="Arquivo:", anchor="w").grid(
            row=2, column=0, padx=20, pady=8, sticky="w"
        )
        arquivo_label = ctk.CTkLabel(
            dialog, text="Nenhum arquivo selecionado",
            text_color="gray60", anchor="w", wraplength=250,
        )
        arquivo_label.grid(row=3, column=0, columnspan=3, padx=20, pady=(0, 4), sticky="w")

        path_holder = {"path": ""}
        btn_confirmar = ctk.CTkButton(dialog, text="Importar", state="disabled", width=140)

        def selecionar():
            path = filedialog.askopenfilename(
                parent=dialog,
                filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
            )
            if not path:
                return
            path_holder["path"] = path
            nome = path.replace("\\", "/").split("/")[-1]
            arquivo_label.configure(text=nome, text_color=("gray10", "gray80"))
            btn_confirmar.configure(state="normal")

        def mostrar_info():
            messagebox.showinfo(
                "Como preparar a planilha",
                "A planilha a ser importada deve ser gerada na Consulta de Eventos "
                "com os seguintes parâmetros:\n\n"
                "• Tipo de Folha: 1 - Mensal\n"
                "• Referência: Mês de referência\n"
                "• Conta: 22, 24, 830 e 831\n\n"
                "Após gerar a consulta, deve-se deixar somente as colunas:\n"
                "CPF · MATRÍCULA · NOME · EVENTO · QTD · VALOR · MÊS ANTERIOR\n\n"
                "Na hora da exportação, selecione o formato XLSX.",
                parent=dialog,
            )

        ctk.CTkButton(dialog, text="📂 Selecionar Arquivo", command=selecionar, width=160).grid(
            row=2, column=1, padx=(0, 6), pady=8, sticky="w"
        )
        ctk.CTkButton(dialog, text="i", command=mostrar_info, width=30).grid(
            row=2, column=2, padx=(0, 20), pady=8, sticky="w"
        )

        def confirmar():
            ano_str = ano_var.get().strip()
            if not ano_str.isdigit() or len(ano_str) != 4:
                messagebox.showwarning("Aviso", "Informe um ano válido (ex: 2026).", parent=dialog)
                return
            self._mes_importacao = MESES.index(mes_var.get()) + 1
            self._ano_importacao = int(ano_str)
            dialog.destroy()
            self._iniciar_importacao(path_holder["path"])

        btn_confirmar.configure(command=confirmar)
        btn_confirmar.grid(row=4, column=0, columnspan=3, pady=(12, 20))

    def _iniciar_importacao(self, path):
        nome = path.replace("\\", "/").split("/")[-1]
        self.label_arquivo.configure(text=f"Carregando {nome}…", text_color="gray60")
        self._set_loading(True)
        threading.Thread(target=self._thread_importar, args=(path,), daemon=True).start()

    def _on_importado(self, path, df):
        faltando = [c for c in self.colunas_necessarias() if c not in df.columns]
        if faltando:
            self._set_loading(False)
            for attr in ("_mes_importacao", "_ano_importacao"):
                if hasattr(self, attr):
                    delattr(self, attr)
            messagebox.showerror(
                "Importação cancelada",
                "A planilha não possui as seguintes colunas obrigatórias:\n\n"
                + "\n".join(f"• {c}" for c in faltando),
            )
            return

        super()._on_importado(path, df)
        mes = getattr(self, "_mes_importacao", None)
        ano = getattr(self, "_ano_importacao", None)
        if mes is not None and ano is not None:
            self.mes_atual = mes
            self.ano_atual = ano
            self._salvar_historico(df, mes, ano)
            del self._mes_importacao
            del self._ano_importacao

    def _salvar_historico(self, df_raw, mes, ano):
        cols_necessarias = ["CPF", "MATRÍCULA", "NOME", "QTD", "VALOR", "ÓRGAO"]
        faltando = [c for c in cols_necessarias if c not in df_raw.columns]
        if faltando:
            messagebox.showerror(
                "Histórico não salvo",
                f"Coluna(s) ausente(s) na planilha: {', '.join(faltando)}\n"
                "O histórico não foi atualizado.",
            )
            return

        df = df_raw.copy()
        for col in ["QTD", "VALOR"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        agg = df.groupby("CPF", as_index=False).agg(
            MATRÍCULA=("MATRÍCULA", "first"),
            NOME=("NOME", "first"),
            ÓRGAO=("ÓRGAO", "first"),
            QTD=("QTD", "sum"),
            VALOR=("VALOR", "sum"),
        )
        registros = agg[["CPF", "MATRÍCULA", "NOME", "QTD", "VALOR", "ÓRGAO"]].to_dict(orient="records")

        historico = {}
        if os.path.exists(_HISTORICO_PATH):
            try:
                with open(_HISTORICO_PATH, "r", encoding="utf-8") as f:
                    historico = json.load(f)
            except Exception:
                historico = {}

        chave = f"{ano}-{mes:02d}"
        if chave in historico:
            if not messagebox.askyesno(
                "Substituir dados",
                f"Já existem dados importados para {MESES[mes - 1]}/{ano}.\nDeseja substituir?",
            ):
                return

        historico[chave] = registros
        try:
            with open(_HISTORICO_PATH, "w", encoding="utf-8") as f:
                json.dump(historico, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._set_status(f"Aviso: não foi possível salvar histórico — {e}", "warning")

    # ------------------------------------------------------------------ Exportar

    def exportar(self):
        if self.df_atual is None:
            messagebox.showwarning("Aviso", "Nenhuma planilha importada!")
            return
        self._abrir_dialog_exportar()

    def _abrir_dialog_exportar(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Exportar")
        dialog.geometry("500x300")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.grid_columnconfigure(1, weight=1)

        # Diretório
        ctk.CTkLabel(dialog, text="Diretório:", anchor="w").grid(
            row=0, column=0, padx=20, pady=(20, 8), sticky="w"
        )
        ultimo_dir = _ler_prefs().get("ultimo_dir_exportar", "")
        dir_var = ctk.StringVar(value=ultimo_dir)
        ctk.CTkEntry(dialog, textvariable=dir_var, width=270).grid(
            row=0, column=1, padx=(0, 5), pady=(20, 8), sticky="ew"
        )

        def escolher_dir():
            d = filedialog.askdirectory(parent=dialog, initialdir=dir_var.get() or None)
            if d:
                dir_var.set(d)

        ctk.CTkButton(dialog, text="📁", width=36, command=escolher_dir).grid(
            row=0, column=2, padx=(0, 20), pady=(20, 8)
        )

        # Nome do arquivo
        ctk.CTkLabel(dialog, text="Nome:", anchor="w").grid(
            row=1, column=0, padx=20, pady=8, sticky="w"
        )
        nome_var = ctk.StringVar()
        ctk.CTkEntry(dialog, textvariable=nome_var, width=270).grid(
            row=1, column=1, padx=(0, 20), pady=8, sticky="ew", columnspan=2
        )

        # Formato
        ctk.CTkLabel(dialog, text="Formato:", anchor="w").grid(
            row=2, column=0, padx=20, pady=8, sticky="w"
        )
        fmt_var = ctk.StringVar(value="Excel")
        ctk.CTkSegmentedButton(
            dialog, values=["Excel", "PDF"], variable=fmt_var, width=160,
        ).grid(row=2, column=1, padx=(0, 20), pady=8, sticky="w", columnspan=2)

        # Toggle abas individuais (Excel only)
        toggle_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        toggle_frame.grid(row=3, column=0, columnspan=3, padx=20, pady=(0, 4), sticky="w")
        abas_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            toggle_frame,
            text="Criar abas individuais para cada secretaria",
            variable=abas_var,
        ).pack(side="left")

        def _nome_padrao():
            mes_nome = MESES[self.mes_atual - 1] if self.mes_atual else ""
            prefixo = f"{mes_nome} {self.ano_atual} " if mes_nome else ""
            sufixo = "horas extras por secretaria" if abas_var.get() else "horas extras geral"
            return f"{prefixo}{sufixo}"

        nome_var.set(_nome_padrao())

        def on_abas(*_):
            nome_var.set(_nome_padrao())

        abas_var.trace_add("write", on_abas)

        def on_fmt(*_):
            if fmt_var.get() == "PDF":
                toggle_frame.grid_remove()
            else:
                toggle_frame.grid()

        fmt_var.trace_add("write", on_fmt)

        def confirmar():
            diretorio = dir_var.get().strip()
            if not diretorio or not os.path.isdir(diretorio):
                messagebox.showwarning("Aviso", "Selecione um diretório válido.", parent=dialog)
                return

            nome = nome_var.get().strip()
            if not nome:
                messagebox.showwarning("Aviso", "Informe o nome do arquivo.", parent=dialog)
                return

            _salvar_pref("ultimo_dir_exportar", diretorio)
            fmt = fmt_var.get()
            ext = ".pdf" if fmt == "PDF" else ".xlsx"
            destino = os.path.join(diretorio, f"{nome}{ext}")

            if os.path.exists(destino):
                if not messagebox.askyesno(
                    "Substituir arquivo",
                    f'O arquivo "{nome}{ext}" já existe.\nDeseja substituir?',
                    parent=dialog,
                ):
                    return

            df_copy = self.df_atual.copy()

            if fmt == "PDF":
                dialog.destroy()
                self._set_loading(True)
                threading.Thread(
                    target=self._thread_exportar_pdf,
                    args=(df_copy, destino, self.mes_atual, self.ano_atual),
                    daemon=True,
                ).start()
            else:
                abas = abas_var.get()
                dialog.destroy()
                self._set_loading(True)
                threading.Thread(
                    target=self._thread_exportar_excel,
                    args=(df_copy, destino, abas),
                    daemon=True,
                ).start()

        ctk.CTkButton(dialog, text="Exportar", command=confirmar, width=140).grid(
            row=4, column=0, columnspan=3, pady=(12, 20)
        )

    # ------------------------------------------------------------------ Excel

    @staticmethod
    def _fmt_br(v):
        try:
            return f"{float(v):.2f}".replace(".", ",")
        except (TypeError, ValueError):
            return str(v)

    def _formatar_df_export(self, df):
        df = df.copy()
        if "TOTAL HORAS MES" in df.columns:
            df["TOTAL HORAS MES"] = df["TOTAL HORAS MES"].apply(self._fmt_br)
        if "DIFERENÇA MES ANTERIOR" in df.columns:
            df["DIFERENÇA MES ANTERIOR"] = df["DIFERENÇA MES ANTERIOR"].apply(
                lambda v: self._celula_formato("DIFERENÇA MES ANTERIOR", v)
            )
        # TOTAL VALOR MES e MES ANTERIOR ficam numéricos — formatados via openpyxl
        return df

    def _criar_df_secretarias(self, df):
        agg = (
            df.groupby("ÓRGAO", as_index=False)
            .agg(
                QTD=("TOTAL HORAS MES", "sum"),
                VALOR=("TOTAL VALOR MES", "sum"),
                MES_ANT=("MES ANTERIOR", "sum"),
            )
            .rename(columns={
                "ÓRGAO": "SECRETARIA",
                "QTD": "QTD GERAL MES",
                "VALOR": "VALOR GERAL MES",
                "MES_ANT": "VALOR MES ANTERIOR",
            })
        )
        if "QTD GERAL MES" in agg.columns:
            agg["QTD GERAL MES"] = agg["QTD GERAL MES"].apply(self._fmt_br)
        # VALOR GERAL MES e VALOR MES ANTERIOR ficam numéricos — formatados via openpyxl
        return agg

    @staticmethod
    def _aplicar_moeda_ws(ws, nomes_cols):
        fmt = 'R$ #,##0.00'
        header = {cell.value: cell.column for cell in ws[1]}
        for nome in nomes_cols:
            idx = header.get(nome)
            if idx is None:
                continue
            for (cell,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                cell.number_format = fmt

    def _thread_exportar_excel(self, df, destino, abas_individuais):
        try:
            df_sec = self._criar_df_secretarias(df)
            df = self._formatar_df_export(df)
            with pd.ExcelWriter(destino, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Geral", index=False)
                df_sec.to_excel(writer, sheet_name="Geral - Secretarias", index=False)
                if abas_individuais:
                    for orgao, grupo in df.groupby("ÓRGAO"):
                        nome = self._nome_aba(str(orgao))
                        grupo.to_excel(writer, sheet_name=nome, index=False)

                wb = writer.book
                self._aplicar_moeda_ws(wb["Geral"], ["TOTAL VALOR MES", "MES ANTERIOR"])
                self._aplicar_moeda_ws(
                    wb["Geral - Secretarias"], ["VALOR GERAL MES", "VALOR MES ANTERIOR"]
                )
                for ws in wb.worksheets:
                    if ws.title not in ("Geral", "Geral - Secretarias"):
                        self._aplicar_moeda_ws(ws, ["TOTAL VALOR MES", "MES ANTERIOR"])
                for ws in wb.worksheets:
                    self._autofit_colunas(ws)
            self.after(0, self._on_exportado)
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self._on_erro_exportar(msg))

    def _autofit_colunas(self, ws):
        for col in ws.columns:
            max_len = max(
                (len(str(cell.value)) for cell in col if cell.value is not None),
                default=8,
            )
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
        ws.auto_filter.ref = ws.dimensions

    def _nome_aba(self, orgao_str):
        nome = re.sub(r'^\d+\s*-\s*', '', orgao_str).strip()
        sem_sec = re.sub(
            r'^SECRETARIA\s+(?:DE\s+|DO\s+|DA\s+|DOS\s+|DAS\s+)?',
            '', nome, flags=re.IGNORECASE,
        ).strip()
        if sem_sec != nome:
            nome = sem_sec.title()
        return re.sub(r'[\\/*?:\[\]]', '', nome)[:31].strip() or "Aba"

    # ------------------------------------------------------------------ PDF

    def _thread_exportar_pdf(self, df, destino, mes, ano):
        try:
            self._gerar_pdf(df, destino, mes, ano)
            self.after(0, self._on_exportado)
        except ImportError:
            self.after(0, lambda: self._on_erro_exportar(
                "A biblioteca fpdf2 não está instalada.\nExecute: pip install fpdf2"
            ))
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self._on_erro_exportar(msg))

    def _gerar_pdf(self, df, destino, mes, ano):
        from fpdf import FPDF

        mes_nome = MESES[mes - 1] if 1 <= mes <= 12 else ""
        periodo = f"{mes_nome}/{ano}" if mes_nome else ""

        # Landscape A4 usable width ≈ 277mm (10mm margins each side)
        cabecs   = ["Matrícula", "Nome",  "Total Horas", "Total Valor"]
        cols_df  = ["MATRÍCULA", "NOME",  "TOTAL HORAS MES", "TOTAL VALOR MES"]
        larguras = [30,           140,     50,                57]
        aligns   = ["C",          "L",     "C",               "R"]

        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_margins(10, 10, 10)

        for orgao, grupo in df.groupby("ÓRGAO"):
            pdf.add_page()
            orgao_nome = re.sub(r'^\d+\s*-\s*', '', str(orgao)).strip()

            # Cabeçalho da página
            pdf.set_font("Helvetica", "B", 13)
            pdf.cell(0, 8, orgao_nome, align="C")
            pdf.ln()
            if periodo:
                pdf.set_font("Helvetica", "", 10)
                pdf.cell(0, 6, periodo, align="C")
                pdf.ln()
            pdf.ln(3)

            # Cabeçalho da tabela
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_fill_color(31, 83, 141)
            pdf.set_text_color(255, 255, 255)
            for cab, larg in zip(cabecs, larguras):
                pdf.cell(larg, 7, cab, border=1, align="C", fill=True)
            pdf.ln()

            # Linhas de dados
            pdf.set_font("Helvetica", "", 9)
            pdf.set_text_color(0, 0, 0)
            for idx, (_, row) in enumerate(grupo.iterrows()):
                fill = idx % 2 == 0
                pdf.set_fill_color(240, 240, 240)
                for col, larg, al in zip(cols_df, larguras, aligns):
                    pdf.cell(larg, 6, self._celula_formato(col, row[col]), border=1, align=al, fill=fill)
                pdf.ln()

        pdf.output(destino)

    # ------------------------------------------------------------------ Validar / Processar

    def _validar_e_exibir(self, df):
        faltando = [c for c in self.colunas_necessarias() if c not in df.columns]
        if faltando:
            self._set_status(
                f"Aviso: coluna(s) não encontrada(s): {', '.join(faltando)}", "warning"
            )
            self._atualizar_preview(df)
            return

        linhas_originais = len(df)
        df_proc = self.processar(df)
        self.df_atual = df_proc
        self._set_status(
            f"Arquivo carregado — {linhas_originais} linha(s) originais → {len(df_proc)} colaborador(es).",
            "success",
        )
        self._atualizar_preview(df_proc)

    def processar(self, df):
        df = df.copy()

        df["EVENTO"] = (
            df["EVENTO"].astype(str).str.extract(r"^(\d+)", expand=False).astype(int)
        )

        for col in ["QTD", "VALOR", "MÊS ANTERIOR"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        agg = (
            df.groupby("CPF", as_index=False)
            .agg(
                MATRÍCULA=("MATRÍCULA", "first"),
                NOME=("NOME", "first"),
                ÓRGAO=("ÓRGAO", "first"),
                QTD=("QTD", "sum"),
                VALOR=("VALOR", "sum"),
                MES_ANT=("MÊS ANTERIOR", "sum"),
            )
            .rename(columns={
                "QTD": "TOTAL HORAS MES",
                "VALOR": "TOTAL VALOR MES",
                "MES_ANT": "MES ANTERIOR",
            })
        )

        agg["TOTAL HORAS MES"] = agg["TOTAL HORAS MES"].round(2)
        agg["TOTAL VALOR MES"] = agg["TOTAL VALOR MES"].round(2)
        agg["MES ANTERIOR"] = agg["MES ANTERIOR"].round(2)

        def _diferenca(row):
            if row["MES ANTERIOR"] == 0:
                return 0.0
            return round(
                (row["TOTAL VALOR MES"] - row["MES ANTERIOR"]) / row["MES ANTERIOR"] * 100, 2
            )

        agg["DIFERENÇA MES ANTERIOR"] = agg.apply(_diferenca, axis=1)

        return agg[[
            "CPF", "MATRÍCULA", "NOME", "ÓRGAO",
            "TOTAL HORAS MES", "TOTAL VALOR MES",
            "MES ANTERIOR", "DIFERENÇA MES ANTERIOR",
        ]]
