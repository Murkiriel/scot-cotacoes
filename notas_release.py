"""Gera as notas do release da última cotação salva (usado pelo GitHub Actions).

Uso:
    python notas_release.py notas.md

Escreve as notas em Markdown no arquivo indicado e imprime, em linhas `chave=valor`,
a tag e o título do release (formato aceito por $GITHUB_OUTPUT). O release segue a
data mais recente do boi gordo.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
from scot_cotacoes import GRUPOS  # noqa: E402
ARQS_CSV = sorted((BASE / 'dados').glob('scot_cotacoes_*.csv'))

# Praças do boi gordo mostradas no resumo (as demais estão nos arquivos).
PRACAS_RESUMO = ['SP Barretos', 'MG Triângulo', 'GO Goiânia', 'MS C. Grande', 'MT Cuiabá',
                 'PA Redenção', 'TO Sul', 'PR Noroeste']
SETAS = {'Subiu': '▲', 'Desceu': '▼', 'Estável': '='}


def br(valor: str, casas: int = 2) -> str:
    """Número no formato brasileiro (1.234,56); vazio vira travessão."""
    if valor in (None, ''):
        return '–'
    texto = f'{float(valor):,.{casas}f}'
    return texto.replace(',', '_').replace('.', ',').replace('_', '.')


def data_br(iso: str) -> str:
    ano, mes, dia = iso[:10].split('-')
    return f'{dia}/{mes}/{ano}'


def main() -> int:
    destino = Path(sys.argv[1])
    linhas = []
    for arquivo in ARQS_CSV:
        with open(arquivo, newline='', encoding='utf-8') as f:
            linhas += list(csv.DictReader(f))
    boi = [l for l in linhas if l['Grupo'] == 'Boi gordo']
    if not boi:
        raise SystemExit('CSV consolidado sem boi gordo')
    ultima = max(l['Data'] for l in boi)

    def valor(grupo, praca, indicador, data=ultima):
        for l in linhas:
            if (l['Grupo'], l['Praça'], l['Indicador'], l['Data']) == (grupo, praca, indicador, data):
                return l
        return None

    partes = [
        f'## Cotações de {data_br(ultima)}',
        '',
        '### Boi gordo (R$/@)',
        '',
        '| Praça | À vista | 30 dias | Indicador Scot | Tendência |',
        '|---|---:|---:|---:|:---:|',
    ]
    for praca in PRACAS_RESUMO:
        vista = valor('Boi gordo', praca, 'À vista (bruto)')
        if not vista:
            continue
        prazo = valor('Boi gordo', praca, '30 dias (bruto)')
        indicador = valor('Indicador do boi gordo', praca, 'Indicador')
        partes.append(f'| {praca} | {br(vista["Valor"])} | {br(prazo["Valor"] if prazo else "")} '
                      f'| {br(indicador["Valor"] if indicador else "")} '
                      f'| {SETAS.get(vista["Tendência"], "")} |')

    partes += ['', '### Datas mais recentes por grupo', '', '| Grupo | Data | Linhas |', '|---|---|---:|']
    grupos = sorted({l['Grupo'] for l in linhas},
                    key=lambda g: GRUPOS.index(g) if g in GRUPOS else len(GRUPOS))
    for grupo in grupos:
        datas = [l['Data'] for l in linhas if l['Grupo'] == grupo]
        recente = max(datas)
        partes.append(f'| {grupo} | {data_br(recente)} | {datas.count(recente)} |')

    datas_boi = sorted({l['Data'] for l in boi})
    partes += [
        '',
        '### Arquivos',
        '',
        '- `scot_cotacoes.xlsx`: planilha com uma aba por grupo e uma aba que explica cada um',
        f'- `scot_cotacoes_AAAA.csv`: histórico no formato longo, um arquivo por ano, {len(linhas)} linhas '
        f'(boi gordo de {data_br(datas_boi[0])} a {data_br(datas_boi[-1])}, '
        f'{len(datas_boi)} data{"s" if len(datas_boi) > 1 else ""})',
    ]
    destino.write_text('\n'.join(partes) + '\n', encoding='utf-8')
    print(f'tag=cotacao-{ultima}')
    print(f'titulo=Cotações {data_br(ultima)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
