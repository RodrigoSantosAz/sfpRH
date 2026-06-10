import datetime
from collections import defaultdict
from dataclasses import dataclass

from .ferias_parser import RegistroFerias, parse_excel  # noqa: F401 (re-exported)


@dataclass
class PessoaVencida:
    matricula: int
    nome: str
    periodos: list[RegistroFerias]


def classificar(
    registros: list[RegistroFerias],
    data_vencimento: datetime.date,
    data_emissao: datetime.date,
) -> list[PessoaVencida]:
    """Retorna, em ordem alfabética, as pessoas acima do limite de férias.

    Critérios para inclusão:
    - 2 ou mais períodos com dias_pendentes > 0, OU
    - 1 período com dias_pendentes > 0 cujo término estimado
      (data_emissao + dias_pendentes dias) ultrapassa data_vencimento.
    """
    validos = [r for r in registros if r.dias_pendentes > 0]

    grupos: dict[tuple, list[RegistroFerias]] = defaultdict(list)
    for r in validos:
        grupos[(r.matricula, r.nome)].append(r)

    resultado: list[PessoaVencida] = []
    for (mat, nome), periodos in grupos.items():
        if len(periodos) >= 2:
            resultado.append(PessoaVencida(mat, nome, periodos))
        elif len(periodos) == 1:
            p = periodos[0]
            fim_estimado = data_emissao + datetime.timedelta(days=p.dias_pendentes)
            if fim_estimado > data_vencimento:
                resultado.append(PessoaVencida(mat, nome, periodos))

    resultado.sort(key=lambda x: x.nome)
    return resultado
