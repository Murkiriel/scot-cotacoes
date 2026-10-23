// Cotações da Scot Consultoria lidas dos CSVs que o GitHub Actions deste repositório
// atualiza todo dia (um arquivo por ano, juntados aqui). Formato longo: uma linha por
// data, praça e indicador. Não depende de Python nem de gateway.
//
// Power BI Desktop > Transformar dados > Nova Fonte > Consulta Nula > Editor Avançado >
// cole este arquivo > Concluído. Credencial: Anônimo. Privacidade: Público.
let
    // Endereço fixo + RelativePath: assim o Power BI Service aceita a atualização agendada.
    Base = "https://raw.githubusercontent.com",
    Pasta = "Murkiriel/scot-cotacoes/main/dados/",
    Anos = {2025 .. Date.Year(DateTime.LocalNow())},
    LerAno = (ano as number) as table =>
        let
            Resposta = Web.Contents(
                Base,
                [
                    RelativePath = Pasta & "scot_cotacoes_" & Text.From(ano) & ".csv",
                    ManualStatusHandling = {404},
                    Timeout = #duration(0, 0, 2, 0)
                ]
            ),
            // Na virada do ano o arquivo novo só aparece na primeira coleta: 404 = ano vazio.
            Tabela =
                if Value.Metadata(Resposta)[Response.Status] = 404 then
                    #table({}, {})
                else
                    Table.PromoteHeaders(
                        Csv.Document(Resposta, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                        [PromoteAllScalars = true]
                    )
        in
            Tabela,
    Juntas = Table.Combine(List.Transform(Anos, LerAno)),
    Tipada = Table.TransformColumnTypes(
        Juntas,
        {
            {"Data", type date}, {"Grupo", type text}, {"Categoria", type text},
            {"Produto", type text}, {"UF", type text}, {"Praça", type text},
            {"Indicador", type text}, {"Unidade", type text}, {"Valor", type number},
            {"Tendência", type text}, {"Coletado em", type datetimezone}, {"Origem", type text}
        },
        "en-US"
    )
in
    Tipada
