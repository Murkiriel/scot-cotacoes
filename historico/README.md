# Histórico importado (dez/2025 a out/2026)

O site da Scot só mostra a cotação mais recente. Antes da primeira coleta deste
repositório (22/10/2026), o histórico existe só no git de dois projetos que já liam a Scot:

| Projeto | O que guardava | Entrou aqui |
|---|---|---|
| [Dom1ng0s/Gado-Scraper](https://github.com/Dom1ng0s/Gado-Scraper) | Boi gordo e novilha, uma versão por coleta (desde 27/12/2025) | Boi gordo: à vista, 30 dias e base SP. Novilha: à vista e 30 dias |
| [Vinyassis01/boiapi](https://github.com/Vinyassis01/boiapi) | Boi gordo e reposição, um arquivo por dia (desde 16/12/2025) | Boi gordo (dias sem o Gado-Scraper) e reposição de machos nelore, 13 UFs |

As linhas importadas têm **Origem = "Importado (Gado-Scraper)"** ou **"Importado (boiapi)"**
e **Coletado em** = hora da coleta original. Não têm os preços descontados (Funrural e
Senar) nem a Tendência, que esses projetos não guardavam.

## A data foi corrigida

Os dois projetos gravam a data da **coleta**. A Scot publica a cotação de cada dia de
madrugada (entre ~22h30 e ~6h25 do dia seguinte), então uma coleta feita no dia E mostra
a tabela do último dia em que a Scot publicou antes de E. O script
[`importar_historico.py`](importar_historico.py) aplica essa regra:

1. **Dia sem publicação:** um dia útil em que nenhum grupo mudou em relação ao dia útil
   anterior (com coleta nos dois dias). Saíram 15: 31/12, 01/01, 16 a 18/02 (Carnaval),
   13/03, 03/04, 21/04, 01/05, 01/06, 04/06, 24/06, 10/07, 07/09 e 12/10.
2. **Data de cada versão:** o último dia útil com publicação antes do dia da coleta.
   Versão igual à anterior é a mesma tabela e não cria data.
3. **Mercado parado:** se a tabela de um grupo ficou igual, mas outro grupo mudou na mesma
   coleta, houve publicação; a data entra com os mesmos valores (boi gordo em 26/01, 17/08
   e 20/08; reposição em 17/08).

## Conferências (o script para sem gravar se alguma falhar)

- **Páginas da Scot no Wayback Machine** (em [`ancoras/`](ancoras/), com a data no título):
  boi gordo de 19/02/2026 e 02/10/2026 (99 valores cada) e reposição de 31/08/2026 e
  02/10/2026 (156 valores cada): todos iguais.
- **Primeira coleta própria** (22/10/2026) contra a versão seguinte do Gado-Scraper: boi
  gordo (99 valores) e novilha (66) iguais.
- **Um projeto contra o outro:** datados separadamente, Gado-Scraper e boiapi dão a mesma
  tabela de boi gordo nas 125 datas em comum.
- Reposição: só arquivos do boiapi com as 13 UFs na ordem esperada e peso de referência
  certo (R$/cab ÷ R$/kg = 375, 300, 240 e 195 kg). A reposição de **fêmeas** do boiapi
  ficou fora: ele lê as linhas da página pela posição e mistura UFs (RJ e outra não batem
  com as âncoras).

## Lacunas

Dias úteis sem dado porque nenhum dos projetos coletou: no boi gordo, 07/01, 27/01, 05/02
e 23/04; na novilha, também boa parte de julho/2026 (o Gado-Scraper parou de 01/07 a
27/07); na reposição, 10 dias. Não há como recuperá-los.

## Rodar de novo

```
python historico/importar_historico.py --sem-gravar   # só confere
python historico/importar_historico.py                # grava (não mexe nas datas de coleta própria)
```

O script clona os dois projetos numa pasta temporária (ou usa clones existentes com
`--gado` e `--boiapi`). Rodar de novo sem mudança nos projetos não altera nenhum arquivo.
