import datetime

from .ferias_logica import PessoaVencida

# Landscape A4: 297mm wide, 15mm margins cada lado → 267mm utilizáveis
_WIDTHS = [18, 82, 67, 37, 30, 33]   # Mat | Nome | Período | Dt.Últ.Gozo | Gozados | Pendentes
_HEADERS = ["Mat.", "Nome", "Período Aquisitivo", "Dt. Últ. Gozo", "Dias Gozados", "Dias Pendentes"]
_ALIGNS = ["C", "L", "C", "C", "C", "C"]


def gerar_pdf(
    pessoas: list[PessoaVencida],
    data_vencimento: datetime.date,
    destino: str,
) -> None:
    from fpdf import FPDF

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.set_margins(15, 15, 15)
    pdf.add_page()

    _cabecalho(pdf, data_vencimento)
    _secao_titulo(pdf, "Funcionarios com mais de 1 periodo vencido")

    if not pessoas:
        pdf.set_font("Helvetica", "I", 10)
        pdf.ln(4)
        pdf.cell(0, 8, "Nenhum servidor acima do limite para o período informado.", align="C")
    else:
        _cabecalho_tabela(pdf)
        _linhas_tabela(pdf, pessoas)

    pdf.output(destino)


# ------------------------------------------------------------------ helpers

def _cabecalho(pdf, data_vencimento: datetime.date) -> None:
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "PREFEITURA DE SAO FRANCISCO DE PAULA", align="C")
    pdf.ln()
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "DEPARTAMENTO DE RECURSOS HUMANOS", align="C")
    pdf.ln()
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, "GESTAO DE RECURSOS HUMANOS", align="C")
    pdf.ln(8)

    pdf.set_font("Helvetica", "I", 10)
    texto = f"Este documento exibe as Ferias com vencimento ate a data de {data_vencimento.strftime('%d/%m/%Y')}."
    pdf.cell(0, 6, texto, align="C")
    pdf.ln(12)


def _secao_titulo(pdf, titulo: str) -> None:
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(31, 83, 141)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 8, f"  {titulo}", fill=True)
    pdf.ln(10)
    pdf.set_text_color(0, 0, 0)


def _cabecalho_tabela(pdf) -> None:
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(31, 83, 141)
    pdf.set_text_color(255, 255, 255)
    for cab, w in zip(_HEADERS, _WIDTHS):
        pdf.cell(w, 7, cab, border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_text_color(0, 0, 0)


def _linhas_tabela(pdf, pessoas: list[PessoaVencida]) -> None:
    pdf.set_font("Helvetica", "", 9)
    fill_toggle = False

    for pessoa in pessoas:
        fill_toggle = not fill_toggle
        r, g, b = (242, 242, 242) if fill_toggle else (255, 255, 255)

        for idx, p in enumerate(pessoa.periodos):
            pdf.set_fill_color(r, g, b)

            mat_str = str(pessoa.matricula) if idx == 0 else ""
            nome_str = pessoa.nome if idx == 0 else ""
            periodo = (
                f"{p.periodo_inicio.strftime('%d/%m/%Y')} a "
                f"{p.periodo_fim.strftime('%d/%m/%Y')}"
            )
            gozo = p.dt_ult_gozo.strftime("%d/%m/%Y") if p.dt_ult_gozo else ""

            valores = [mat_str, nome_str, periodo, gozo, str(p.dias_gozados), str(p.dias_pendentes)]
            for val, w, al in zip(valores, _WIDTHS, _ALIGNS):
                pdf.cell(w, 6, val, border=1, align=al, fill=True)
            pdf.ln()
