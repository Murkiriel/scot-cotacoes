"""Importa o histórico de dez/2025 a out/2026 de dois projetos que já liam a Scot.

O site da Scot só mostra a cotação mais recente; antes da primeira coleta deste
repositório (22/10/2026), o histórico existe só no git destes projetos:

    Dom1ng0s/Gado-Scraper   boi gordo e novilha (uma versão por coleta, desde 27/12/2025)
    Vinyassis01/boiapi      boi gordo e reposição (um arquivo por dia, desde 16/12/2025)

Os dois gravam a data da COLETA, não a da tabela. A Scot publica a cotação do dia de
madrugada (entre ~22h30 e ~6h25 do dia seguinte), então uma coleta feita no dia E
mostra a tabela do último dia útil antes de E. Este script:

    1. lê todas as versões dos dois projetos, com o horário de cada coleta;
    2. dá a cada versão a data do dia útil anterior à coleta; versão igual à anterior
       é a mesma tabela (fim de semana, feriado ou dia sem publicação) e não cria data;
    3. confere as datas e os valores com páginas da Scot guardadas pelo Wayback Machine
       (historico/ancoras) e com a primeira coleta deste repositório; se algo não bater,
       para sem gravar nada;
    4. grava as datas que ainda não existem, com Origem = "Importado (<projeto>)".

Entram só os campos que esses projetos guardavam: boi gordo (à vista, 30 dias e base),
novilha (à vista e 30 dias) e reposição de machos nelore (13 UFs, R$/cab, R$/kg e
troca). A reposição de fêmeas do boiapi fica fora: ele lê as linhas pela posição e
mistura UFs (conferido com as âncoras).

Uso (na pasta do repositório):
    python historico/importar_historico.py              # clona os projetos numa pasta temporária
    python historico/importar_historico.py --gado <pasta> --boiapi <pasta>   # clones existentes
    python historico/importar_historico.py --sem-gravar # só confere e mostra o resumo
"""

from __future__ import annotations

import argparse
import collections
import datetime as dt
import json
import re
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path

import lxml.html

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
import scot_cotacoes as sc  # noqa: E402

DIR_ANCORAS = Path(__file__).resolve().parent / 'ancoras'
REPOS = {'Gado-Scraper': 'https://github.com/Dom1ng0s/Gado-Scraper.git',
         'boiapi': 'https://github.com/Vinyassis01/boiapi.git'}
ORIGEM = {'Gado-Scraper': 'Importado (Gado-Scraper)', 'boiapi': 'Importado (boiapi)'}
# Reposição de machos nelore no boiapi: animal -> (início do nome na Scot, peso de referência em kg).
ANIMAIS_BOIAPI = {'boi magro': ('Boi magro', 375), 'garrote': ('Garrote', 300),
                  'bezerro': ('Bezerro', 240), 'bezerro desmamado': ('Desmama', 195)}
UFS_BOIAPI = ['SP', 'MG', 'GO', 'MS', 'BA', 'MT', 'PR', 'PA', 'RO', 'TO', 'AC', 'MA', 'RJ']
MINIMO_PRACAS = 30


def log(msg: str) -> None:
    print(msg, flush=True)


def git(repo: Path, *args: str) -> str:
    return subprocess.run(['git', '-C', str(repo), *args], capture_output=True, check=True).stdout.decode('utf-8')


def clonar(destino: Path, url: str) -> Path:
    # --no-checkout: o boiapi tem nomes de arquivo longos demais para o Windows; só o git é lido.
    subprocess.run(['git', 'clone', '--quiet', '--no-checkout', url, str(destino)], check=True)
    return destino


def chave(texto: str) -> str:
    """Nome da praça sem acentos, espaços, pontos, asteriscos e "(kg)" (os projetos escrevem diferente)."""
    texto = unicodedata.normalize('NFKD', texto.replace('(kg)', '').replace('*', ''))
    return re.sub(r'[^a-z]', '', texto.encode('ascii', 'ignore').decode().lower())


def num(valor) -> float | None:
    if valor in (None, ''):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    return float(texto.replace('.', '').replace(',', '.')) if ',' in texto else float(texto)


def dia_util_anterior(dia: dt.date) -> dt.date:
    dia -= dt.timedelta(days=1)
    while dia.weekday() >= 5:
        dia -= dt.timedelta(days=1)
    return dia


# --------------------------------------------------------------------------- leitura dos projetos

def fotos_gado(repo: Path) -> dict[str, list]:
    """{grupo: [(quando, {chave_praça: {indicador: valor}})]} de cada versão no git."""
    fotos = {'Boi gordo': [], 'Novilha gorda': []}
    campos = {'Boi gordo': [('preco_vista', 'À vista (bruto)'), ('preco_30d', '30 dias (bruto)'),
                            ('variacao', 'Diferencial de base SP')],
              'Novilha gorda': [('preco_vista', 'À vista (bruto)'), ('preco_30d', '30 dias (bruto)')]}
    for arquivo, grupo in (('cotacoes_boi_hoje.json', 'Boi gordo'), ('cotacoes_novilha_hoje.json', 'Novilha gorda')):
        for linha in git(repo, 'log', '--format=%H %aI', '--', arquivo).splitlines():
            commit, quando = linha.split()
            try:
                dados = json.loads(git(repo, 'show', f'{commit}:{arquivo}'))
            except (subprocess.CalledProcessError, json.JSONDecodeError):
                continue
            foto = {}
            for r in dados:
                if isinstance(r.get('praca'), str) and num(r.get('preco_vista')) is not None:
                    foto[chave(r['praca'])] = {ind: num(r.get(c)) for c, ind in campos[grupo]}
            hora = dt.datetime.fromisoformat(quando).astimezone(sc.FUSO_BRASILIA)
            fotos[grupo].append((hora, foto, 'Gado-Scraper'))
    return fotos


def fotos_boiapi(repo: Path) -> dict[str, list]:
    fotos = {'Boi gordo': [], 'Reposição': []}
    nomes = git(repo, 'ls-tree', '--name-only', 'HEAD', 'scrapers/').splitlines()
    for nome in sorted(nomes):
        m = re.search(r'valores_gerais_(boi_gordo|reposicao)_(\d{4}-\d{2}-\d{2})\.json$', nome)
        if not m:
            continue
        linhas = []
        for texto in git(repo, 'show', f'HEAD:{nome}').splitlines():
            try:
                linhas.append(json.loads(texto))
            except json.JSONDecodeError:
                pass  # o arquivo de 20/02/2026 tem linhas com aspas quebradas
        commits = git(repo, 'log', '--format=%aI', '--', nome).split()
        dia = dt.date.fromisoformat(m.group(2))
        horas = sorted(dt.datetime.fromisoformat(c).astimezone(sc.FUSO_BRASILIA) for c in commits)
        hora = next((h for h in horas if h.date() == dia), dt.datetime.combine(dia, dt.time(12), sc.FUSO_BRASILIA))
        foto = {}
        if m.group(1) == 'boi_gordo':
            for r in linhas:
                foto[chave(r['estado'] + r.get('regiao', ''))] = {
                    'À vista (bruto)': num(r['arroba_a_vista']), '30 dias (bruto)': num(r['arroba_a_prazo']),
                    'Diferencial de base SP': num(r['variacao'])}
            fotos['Boi gordo'].append((hora, foto, 'boiapi'))
        else:
            machos = [r for r in linhas if r['animal'] in ANIMAIS_BOIAPI]
            # Só arquivos com a sequência esperada de UFs em cada animal (o leitor do boiapi é por posição).
            if all([r['estado'] for r in machos if r['animal'] == a] == UFS_BOIAPI for a in ANIMAIS_BOIAPI):
                for r in machos:
                    nome_animal, peso = ANIMAIS_BOIAPI[r['animal']]
                    cab, kg = num(r['valor_animal']), num(r['valor_kg'])
                    if cab and kg and abs(cab / kg / peso - 1) < 0.01:
                        foto[(r['estado'], nome_animal)] = {'R$/cab': cab, 'R$/kg': kg,
                                                            'Relação de troca': num(r['relacao_troca'])}
                if len(foto) == len(UFS_BOIAPI) * len(ANIMAIS_BOIAPI):
                    fotos['Reposição'].append((hora, foto, 'boiapi'))
    return fotos


# --------------------------------------------------------------------------- datas

def _validas(fotos: list) -> list:
    """Versões em ordem de coleta, sem arquivos vazios ou incompletos."""
    return [f for f in sorted(fotos, key=lambda f: f[0]) if len(f[1]) >= MINIMO_PRACAS]


def evidencias(fotos: dict) -> dict[dt.date, list[bool]]:
    """Para cada dia útil D: em cada grupo coletado logo depois de D, sem buraco em relação à
    coleta anterior (anterior datada no dia útil antes de D), a tabela mudou (True) ou não."""
    saida = collections.defaultdict(list)
    for lista in fotos.values():
        anterior = None  # (data, foto) da última data vista
        for quando, foto, _ in _validas(lista):
            data = dia_util_anterior(quando.date())
            if anterior and data == anterior[0]:
                continue
            if anterior and anterior[0] == dia_util_anterior(data):
                saida[data].append(foto != anterior[1])
            anterior = (data, foto)
    return saida


def ultimo_dia_publicado(sem_publicacao: set[dt.date]):
    """Data da tabela que uma coleta no dia E mostra: o último dia útil antes de E em que a Scot publicou."""
    def rotulo(dia: dt.date) -> dt.date:
        data = dia_util_anterior(dia)
        while data in sem_publicacao:
            data = dia_util_anterior(data)
        return data
    return rotulo


def datar(fotos: list, rotulo=dia_util_anterior) -> tuple[dict, list[str]]:
    """{data: (foto, quando, fonte)} a partir das versões em ordem de coleta, e as ocorrências."""
    publicacoes, ocorrencias, anterior = {}, [], None
    for quando, foto, fonte in _validas(fotos):
        assinatura = sorted(foto.items())
        if assinatura == anterior:
            continue  # mesma tabela da versão anterior: não houve publicação nova
        anterior = assinatura
        data = rotulo(quando.date())
        ultima = max(publicacoes) if publicacoes else None
        if ultima and data < ultima:
            ocorrencias.append(f'{quando:%d/%m %H:%M} ({fonte}): tabela nova datada antes da última ({data}); ignorada')
            continue
        if ultima == data:
            ocorrencias.append(f'{data}: tabela mudou de novo em {quando:%d/%m %H:%M} ({fonte}); fica a última')
            publicacoes[data] = (foto, publicacoes[data][1], fonte)  # mantém a hora da primeira coleta
            continue
        publicacoes[data] = (foto, quando, fonte)
    return publicacoes, ocorrencias


def preencher_sem_mudanca(datas: dict, fotos: dict, publicados: set[dt.date], rotulo) -> list[str]:
    """Tabela igual à anterior pode ser dia sem publicação (feriado) ou mercado parado.
    Se outro grupo mudou naquele dia (publicação comprovada), o grupo parado ganha a
    data, com os valores da anterior. Devolve o que foi preenchido."""
    preenchidos = []
    for grupo, por_data in datas.items():
        coletas = {}
        for quando, foto, fonte in _validas(fotos[grupo]):
            coletas.setdefault(rotulo(quando.date()), (quando, fonte))
        for dia in sorted(publicados - set(por_data)):
            anteriores = [d for d in por_data if d < dia]
            if dia in coletas and anteriores:
                quando, fonte = coletas[dia]
                por_data[dia] = (por_data[max(anteriores)][0], quando, fonte)
                preenchidos.append(f'{grupo} {dia:%d/%m/%Y}')
    return preenchidos


# --------------------------------------------------------------------------- conferências

def ancoras() -> dict[tuple[str, str], dict]:
    """{(grupo, data): {chave: {indicador: valor}}} das páginas guardadas pelo Wayback."""
    saida = {}
    for arquivo in sorted(DIR_ANCORAS.glob('*.html')):
        pagina = arquivo.name.split('_')[0]
        doc = lxml.html.fromstring(arquivo.read_text(encoding='utf-8'))
        for r in sc.PAGINAS[pagina](doc):
            if r['grupo'] == 'Boi gordo':
                k = chave(r['praca'])
            elif r['grupo'] == 'Reposição' and r['categoria'] == 'Macho nelore' and '(' not in r['praca']:
                k = (r['uf'], r['produto'].split(' (')[0])
            else:
                continue
            saida.setdefault((r['grupo'], r['data']), {}).setdefault(k, {})[r['indicador']] = r['valor']
    return saida


def conferir(grupo: str, datas: dict, referencia: dict, nome: str) -> int:
    """Compara as datas importadas com uma referência {(grupo, data): foto}; devolve o nº de divergências."""
    erros = 0
    for (g, data), foto_ref in referencia.items():
        if g != grupo:
            continue
        dia = dt.date.fromisoformat(data)
        if dia not in datas:
            log(f'  {nome} {grupo} {data}: data não encontrada no histórico importado')
            erros += 1
            continue
        foto = datas[dia][0]
        diferencas = [(k, i) for k, inds in foto.items() for i, v in inds.items()
                      if k in foto_ref and i in foto_ref[k] and foto_ref[k][i] != v]
        comuns = sum(1 for k, inds in foto.items() for i in inds if k in foto_ref and i in foto_ref[k])
        situacao = 'OK' if not diferencas and comuns else f'{len(diferencas)} diferenças'
        log(f'  {nome} {grupo} {data}: {comuns} valores comparados, {situacao}')
        erros += bool(diferencas) or not comuns
    return erros


# --------------------------------------------------------------------------- execução

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    ap.add_argument('--gado', type=Path, help='clone existente de Dom1ng0s/Gado-Scraper')
    ap.add_argument('--boiapi', type=Path, help='clone existente de Vinyassis01/boiapi')
    ap.add_argument('--sem-gravar', action='store_true', help='só confere e mostra o resumo')
    args = ap.parse_args()

    temporaria = Path(tempfile.mkdtemp(prefix='scot_historico_'))
    gado = args.gado or clonar(temporaria / 'Gado-Scraper', REPOS['Gado-Scraper'])
    boiapi = args.boiapi or clonar(temporaria / 'boiapi', REPOS['boiapi'])
    fotos = collections.defaultdict(list)
    for origem in (fotos_gado(gado), fotos_boiapi(boiapi)):
        for grupo, lista in origem.items():
            fotos[grupo] += lista
    for grupo, lista in fotos.items():
        log(f'{grupo}: {len(lista)} versões lidas '
            f'({", ".join(f"{f} {n}" for f, n in collections.Counter(x[2] for x in lista).items())})')

    # Nomes, UFs e unidades como a Scot escreve, tirados da coleta própria mais recente.
    proprias = [l for l in sc.ler_todos_os_dias() if l['origem'] == sc.ORIGEM_PROPRIA]
    _exigir = sc._exigir
    _exigir(bool(proprias), 'rode antes uma coleta própria (scot_cotacoes.py): ela dá os nomes das praças')
    nomes = {}
    for l in proprias:
        if l['grupo'] in ('Boi gordo', 'Novilha gorda') and l['indicador'] == 'À vista (bruto)':
            nomes[(l['grupo'], chave(l['praca']))] = (l['praca'], l['uf'], l['unidade'])
        if l['grupo'] == 'Reposição' and l['categoria'] == 'Macho nelore' and '(' not in l['praca']:
            nomes[('Reposição', (l['uf'], l['produto'].split(' (')[0]))] = (l['praca'], l['uf'], l['produto'])
    datas_proprias = {(l['grupo'], l['data']) for l in proprias}

    # Datas por grupo; no boi gordo, cada projeto também é datado sozinho para conferir um com o outro.
    # Dias sem publicação (feriados): nenhum grupo mudou. Dias com publicação comprovada: algum mudou.
    ev = evidencias(fotos)
    sem_publicacao = {d for d, mudou in ev.items() if not any(mudou)}
    publicados = {d for d, mudou in ev.items() if any(mudou)}
    rotulo = ultimo_dia_publicado(sem_publicacao)
    log(f'Dias úteis sem publicação da Scot (nenhum grupo mudou): '
        + ' '.join(f'{d:%d/%m/%y}' for d in sorted(sem_publicacao)))

    datas, erros = {}, 0
    for grupo, lista in fotos.items():
        datas[grupo], ocorrencias = datar(lista, rotulo)
        for o in ocorrencias:
            log(f'  {grupo}: {o}')
    for p in preencher_sem_mudanca(datas, fotos, publicados, rotulo):
        log(f'  {p}: tabela igual à anterior, mas outro grupo mudou na mesma coleta; mantida com os mesmos valores')
    por_fonte = {f: datar([x for x in fotos['Boi gordo'] if x[2] == f], rotulo)[0] for f in REPOS}
    comuns = set(por_fonte['Gado-Scraper']) & set(por_fonte['boiapi'])
    divergentes = [d for d in comuns if por_fonte['Gado-Scraper'][d][0] != por_fonte['boiapi'][d][0]]
    log(f'Boi gordo, Gado-Scraper x boiapi datados separadamente: {len(comuns)} datas em comum, '
        f'{len(divergentes)} com valores diferentes {sorted(map(str, divergentes))[:5]}')
    erros += len(divergentes)

    log('Conferência com as páginas do Wayback Machine (data no título da tabela):')
    ref = ancoras()
    erros += sum(conferir(g, datas[g], ref, 'Wayback') for g in ('Boi gordo', 'Reposição'))
    log('Conferência com a primeira coleta própria (a versão seguinte a ela nos projetos):')
    propria = {}
    for l in proprias:
        if l['grupo'] in ('Boi gordo', 'Novilha gorda') and l['indicador'] in (
                'À vista (bruto)', '30 dias (bruto)', 'Diferencial de base SP'):
            propria.setdefault((l['grupo'], l['data']), {}).setdefault(chave(l['praca']), {})[l['indicador']] = float(l['valor'])
    erros += sum(conferir(g, datas[g], {k: v for k, v in propria.items() if k[0] == g}, 'Coleta própria')
                 for g in ('Boi gordo', 'Novilha gorda'))
    if erros:
        log(f'{erros} divergência(s): nada foi gravado')
        return 1

    registros, faltando = [], collections.Counter()
    for grupo, por_data in datas.items():
        for data, (foto, quando, fonte) in sorted(por_data.items()):
            if (grupo, data.isoformat()) in datas_proprias:
                continue
            for k, indicadores in foto.items():
                if (grupo, k) not in nomes:
                    faltando[(grupo, str(k))] += 1
                    continue
                praca, uf, extra = nomes[(grupo, k)]
                for indicador, valor in indicadores.items():
                    if valor is None:
                        continue
                    if grupo == 'Reposição':
                        unidade = {'R$/cab': 'R$/cab', 'R$/kg': 'R$/kg', 'Relação de troca': '@ de boi gordo'}[indicador]
                        r = sc._registro(data.isoformat(), grupo, indicador, valor, unidade,
                                         categoria='Macho nelore', produto=extra, uf=uf, praca=praca)
                    else:
                        unidade = '%' if indicador == 'Diferencial de base SP' else extra
                        r = sc._registro(data.isoformat(), grupo, indicador, valor, unidade,
                                         produto=grupo, uf=uf, praca=praca)
                    registros.append({**r, 'coletado_em': quando.isoformat(), 'origem': ORIGEM[fonte]})
    for (grupo, k), n in faltando.items():
        log(f'  AVISO {grupo}: praça "{k}" não existe na coleta própria; {n} data(s) ignoradas')

    resumo = collections.Counter((r['grupo'], r['data']) for r in registros)
    for grupo in datas:
        ds = sorted(d for g, d in resumo if g == grupo)
        if ds:
            log(f'{grupo}: {len(ds)} datas, de {ds[0]} a {ds[-1]}, {sum(n for (g, _), n in resumo.items() if g == grupo)} linhas')
    if args.sem_gravar:
        log('Conferência concluída (--sem-gravar: nada gravado)')
        return 0
    controle = sc.carregar_controle()
    mudancas = sc.gravar(registros, controle, sc.agora().isoformat())
    linhas = sc.ler_todos_os_dias()
    log(f'{len(mudancas)} grupo(s)/data(s) gravados; CSV: {sc.gerar_csv(linhas)} ({len(linhas)} linhas)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
