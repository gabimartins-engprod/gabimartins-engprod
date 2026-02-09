// =====================================================================================
// CONSULTA: Monitoramento_Extracao
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Visão histórica por ID baseada no Histórico_Bitrix, permitindo identificar
//             primeira e última aparição, bem como o status Dentro/Fora da extração atual.
// ORIGEM DOS DADOS: Histórico_Bitrix + Registros_Contatos_Auto
// SAÍDA: ID | Primeira Aparição | Última Aparição | Status | Data Saída | NomeArquivo (Último)
// VERSÃO: Atualizado em 09/02/2026 
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES:
// - Consulta de apoio à validação técnica.
// - Não representa rotina operacional.
// – ajuste semântico (Primeira Aparição)
// =====================================================================================
let
    HistFonte = #"Histórico_Bitrix",

    Hist =
        Table.TransformColumnTypes(
            HistFonte,
            {{"ID", type text}, {"DataArquivo", type date}, {"NomeArquivo", type text}},
            "pt-BR"
        ),

    HistLimpo =
        Table.SelectRows(
            Hist,
            each [ID] <> null and Text.Trim([ID]) <> "" and [DataArquivo] <> null
        ),

    Agrupado =
        Table.Group(
            HistLimpo,
            {"ID"},
            {
                {"Primeira Aparição", each List.Min([DataArquivo]), type date},
                {"Última Aparição", each List.Max([DataArquivo]), type date},
                {"NomeArquivo (Último)", each
                    let
                        t = Table.Sort(_, {{"DataArquivo", Order.Descending}}),
                        na = try t{0}[NomeArquivo] otherwise null
                    in
                        na,
                 type text}
            }
        ),

    AtualIDs =
        Table.Distinct(
            Table.TransformColumnTypes(
                Table.SelectColumns(Registros_Contatos_Auto, {"ID"}),
                {{"ID", type text}},
                "pt-BR"
            )
        ),

    AddStatus =
        Table.AddColumn(
            Agrupado,
            "Status",
            each if Table.Contains(AtualIDs, [ID = [ID]]) then "Dentro da extração" else "Fora da extração",
            type text
        ),

    AddDataSaida =
        Table.AddColumn(
            AddStatus,
            "Data Saída",
            each if [Status] = "Fora da extração" then [Última Aparição] else null,
            type date
        ),

    Ordenado =
        Table.Sort(AddDataSaida, {{"ID", Order.Ascending}})
in
    Ordenado
