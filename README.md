# Cotações da Scot Consultoria

Histórico das cotações publicadas pela [Scot Consultoria](https://www.scotconsultoria.com.br/cotacoes/):
boi gordo, boi China, vaca e novilha gordas, indicador do boi, reposição, mercado futuro,
milho, soja, atacado, equivalentes, boi no mundo, couro e sebo e leite. Coletadas pelo
GitHub Actions quatro vezes por dia.

O site da Scot só mostra a cotação mais recente; este repositório guarda cada data desde
22/10/2026 (primeira coleta). Antes disso há um **histórico importado** de dois projetos
que já liam a Scot, com a data corrigida e conferida: boi gordo e novilha desde 26/12/2025
e reposição de machos nelore desde 23/02/2026 (coluna Origem; ver
[`historico/`](historico/README.md)).

**Baixar:** cada data do boi gordo vira um [release](../../releases) com a planilha e os
CSVs do histórico. Link fixo para a planilha mais recente:
[`scot_cotacoes.xlsx`](../../releases/latest/download/scot_cotacoes.xlsx).

O histórico fica no repositório em **um CSV por ano** (cerca de 50 MB por ano; o GitHub não
aceita arquivo acima de 100 MB): [`dados/scot_cotacoes_2025.csv`](dados/scot_cotacoes_2025.csv),
[`dados/scot_cotacoes_2026.csv`](dados/scot_cotacoes_2026.csv) e assim por diante. O Power BI
lê e junta os anos (ver abaixo); o endereço de cada um é

```
https://raw.githubusercontent.com/Murkiriel/scot-cotacoes/main/dados/scot_cotacoes_AAAA.csv
```

## O que tem no CSV

Formato longo: uma linha por data, praça e indicador, com as colunas **Data, Grupo,
Categoria, Produto, UF, Praça, Indicador, Unidade, Valor, Tendência, Coletado em e Origem**
(Coleta própria ou Importado).
UTF-8, separador vírgula, ponto decimal, datas `AAAA-MM-DD`.

| Grupo | Praças / itens | Indicadores |
|---|---|---|
| Boi gordo | ~33 praças | À vista e 30 dias: bruto, desc. Funrural, desc. Senar; diferencial de base SP (%) |
| Boi China | 11 estados | Bruto e líquido a 30 dias |
| Vaca gorda, Novilha gorda | ~33 praças | Como o boi gordo, sem a base |
| Indicador do boi gordo | ~33 praças | Indicador (média móvel de 5 dias úteis) e variação |
| Reposição | UFs × macho/fêmea nelore e mestiço, macho cruzamento industrial | R$/cab, R$/kg, relação de troca (@ de boi gordo) |
| Mercado futuro do boi | Vencimentos da B3 | Ajuste, ajuste anterior, variação, contratos negociados, câmbio, US$/@ |
| Milho, Soja | ~34 praças | Preço de compra, R$/saca 60 kg |
| Atacado | Cortes, SP | R$/kg |
| Equivalentes | Físico, Scot, carcaça, desossa | Valor (R$/@) e margem (%) |
| Boi no mundo | 8 países | US$/@ |
| Couro e sebo | Brasil Central, RS | R$/kg (semanal) |
| Leite | ~34 regiões | Mínimo e máximo (qualidade e padrão) e médias, R$/L e US$/L (mensal) |

- **Data** é a da cotação publicada pela Scot, não a da coleta. No mercado futuro é a do
  último ajuste da tabela; no leite, o 1º dia do mês do pagamento.
- **Unidade:** no RS o boi, a vaca e a novilha são cotados em R$/kg vivo (50% de
  rendimento de carcaça); nas demais praças, em R$/@.
- **Tendência** é a seta que a Scot mostra na tabela (Estável, Subiu, Desceu), onde houver.
- A planilha tem uma aba por grupo, com uma coluna por indicador, e uma aba **Sobre** com
  as notas da Scot sobre cada tabela.

## Usar no Power BI

Cole [`powerbi/ScotCotacoes_GitHub.pq`](powerbi/ScotCotacoes_GitHub.pq) em
**Transformar dados → Nova Fonte → Consulta Nula → Editor Avançado**. A credencial é
**Anônimo** e a privacidade, **Público**. Para a atualização agendada no Power BI Service
não é preciso gateway. A consulta junta os arquivos de 2025 até o ano atual (ano ainda sem
arquivo, na virada, entra vazio).

Para ter uma coluna por indicador (ex.: à vista e 30 dias lado a lado), filtre o Grupo e
use **Transformar → Coluna Dinâmica** na coluna Indicador, com Valor como valores
(no Leite, junte antes Indicador e Unidade, porque a média vem em R$/L e em US$/L).

## Rodar no próprio computador

Requer Python 3.10+:

```
pip install -r requirements.txt
python scot_cotacoes.py
```

Além do CSV, gera `scot_cotacoes.xlsx`. Opções: `--paginas boi-gordo reposicao ...`,
`--sem-excel`, `--html <pasta>` (lê páginas salvas em vez de baixar) e `--sem-gravar`.
Para testar o leitor com as páginas de 22/10/2026 guardadas em `testes/html`:

```
python scot_cotacoes.py --html testes/html --sem-gravar
```

## Observações

- **Mercado futuro:** a coluna que a Scot chama de "C. Abertos" traz os contratos
  **negociados** no dia (confere com o boletim da B3), por isso aqui ela se chama
  "Contratos negociados". Para os futuros completos da B3, veja
  [b3-futuros-agro](https://github.com/Murkiriel/b3-futuros-agro).
- **Indicador do boi:** a Scot mostra a variação sem sinal (o sinal está na seta); aqui
  ela é calculada como indicador do dia menos o do dia anterior.
