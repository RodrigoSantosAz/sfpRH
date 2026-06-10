"""Parser para o relatório de férias vencidas exportado pelo ADMRH.

O arquivo Excel gerado pelo sistema tem estrutura irregular:
- Linhas 1-5: cabeçalho com título, data e página
- Linha 6: rótulos das colunas
- Linhas de dados: col A numérica (matrícula)
- Rodapé: 'Total Geral de Servidores:' e nome do template

Colunas relevantes (índice 0-based):
  0  → Matrícula
  1  → Nome
  3  → Início do período aquisitivo (datetime)
  5  → Fim do período aquisitivo (datetime)
  6  → Data do último gozo (datetime ou None)
  7  → Dias gozados (int)
  11 → Dias pendentes (int)
"""

import datetime
from dataclasses import dataclass
from typing import Optional


@dataclass
class RegistroFerias:
    matricula: int
    nome: str
    periodo_inicio: datetime.date
    periodo_fim: datetime.date
    dt_ult_gozo: Optional[datetime.date]
    dias_gozados: int
    dias_pendentes: int


def _to_date(value) -> Optional[datetime.date]:
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return None


def parse_excel(caminho: str) -> list[RegistroFerias]:
    """Lê o arquivo .xlsx do ADMRH e retorna os registros de férias."""
    try:
        import openpyxl
    except ImportError as e:
        raise ImportError("openpyxl é necessário para leitura do arquivo.") from e

    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active

    registros = []
    for row in ws.iter_rows(values_only=True):
        matricula_raw = row[0]
        if not isinstance(matricula_raw, (int, float)):
            continue

        matricula = int(matricula_raw)
        nome = str(row[1]).strip() if row[1] else ""
        periodo_inicio = _to_date(row[3])
        periodo_fim = _to_date(row[5])
        dt_ult_gozo = _to_date(row[6])
        dias_gozados = int(row[7]) if row[7] is not None else 0
        dias_pendentes = int(row[11]) if row[11] is not None else 0

        if not nome or periodo_inicio is None or periodo_fim is None:
            continue

        registros.append(RegistroFerias(
            matricula=matricula,
            nome=nome,
            periodo_inicio=periodo_inicio,
            periodo_fim=periodo_fim,
            dt_ult_gozo=dt_ult_gozo,
            dias_gozados=dias_gozados,
            dias_pendentes=dias_pendentes,
        ))

    return registros
