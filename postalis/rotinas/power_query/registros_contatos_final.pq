// =====================================================================================
// CONSULTA: Registros_Contatos_Final
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Consolidar a base automática (Registros_Contatos_Auto) com os dados
//             manuais (Registros_Contatos_Manual), gerando a visão final de contatos.
//             travado em Data_1..Data_5 + proteção Manual 1 por ID + visão enxuta
// ORIGEM DOS DADOS: Registros_Contatos_Auto + Registros_Contatos_Manual
// DESTINO: Tabela "Registros Contatos" (Registros_Contatos_Final) no Excel
// VERSÃO: Atualizado em 09/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// =====================================================================================
let
    // --------------------------------------------------
    // 1) FONTE AUTOMÁTICA: REGISTROS_CONTATOS_AUTO
    // --------------------------------------------------
    Fonte =
        Table.TransformColumnTypes(
            Registros_Contatos_Auto,
            {
                {"ID", type text},
                {"Parent task ID", type text},
                {"Created on", type datetime},
                {"Plano", type text},
                {"Nome", type text},
                {"Matrícula", type text}
            },
            "pt-BR"
        ),

    // --------------------------------------------------
    // 2) BASE MANUAL: GARANTIR 1 LINHA POR ID (PROTEÇÃO)
    // --------------------------------------------------
    Manual_Tipada = Registros_Contatos_Manual,
    Manual_Ordenada = try Table.Sort(Manual_Tipada, {{"Created on", Order.Descending}}) otherwise Manual_Tipada,
    Manual_1porID = Table.Distinct(Manual_Ordenada, {"ID"}),

    ManualNames = Table.ColumnNames(Manual_1porID),

    // --------------------------------------------------
    // 3) ÍNDICES FIXOS (REGRA OFICIAL): SOMENTE 1..5
    //    (e só usa o que existir na tabela)
    // --------------------------------------------------
    IndicesFixos = {1..5},
    IndicesValidos =
        List.Select(
            IndicesFixos,
            (i) => List.Contains(ManualNames, "Data_" & Text.From(i))
        ),

    ColsFixas = {"Observação Extra","E-mail","Telefone"},
    ColsDinamicas =
        List.Combine(
            List.Transform(
                IndicesValidos,
                (i) => {"Data_" & Text.From(i), "OBS_" & Text.From(i), "Responsável_" & Text.From(i)}
            )
        ),
    ManualToExpand = List.Intersect({List.Combine({ColsFixas, ColsDinamicas}), ManualNames}),

    // --------------------------------------------------
    // 4) MERGE COM A TABELA MANUAL (POR ID)
    // --------------------------------------------------
    MergeManual =
        Table.NestedJoin(
            Fonte,
            {"ID"},
            Manual_1porID,
            {"ID"},
            "Manual",
            JoinKind.LeftOuter
        ),

    ExpManual =
        if List.Count(ManualToExpand) = 0 then
            MergeManual
        else
            Table.ExpandTableColumn(MergeManual, "Manual", ManualToExpand, ManualToExpand),

    // --------------------------------------------------
    // 5) ÚLTIMO CONTATO (DATA/OBS/RESP) COERENTES
    // --------------------------------------------------
    AddUltimoContato =
        Table.AddColumn(
            ExpManual,
            "UltimoContato",
            (r as record) as record =>
                let
                    tentativas =
                        if List.Count(IndicesValidos) = 0 then
                            {}
                        else
                            List.Transform(
                                IndicesValidos,
                                (i) =>
                                    let
                                        d = try Date.From(Record.FieldOrDefault(r, "Data_" & Text.From(i), null)) otherwise null,
                                        o = try Record.FieldOrDefault(r, "OBS_" & Text.From(i), null) otherwise null,
                                        resp = try Record.FieldOrDefault(r, "Responsável_" & Text.From(i), null) otherwise null
                                    in
                                        [Idx = i, Data = d, OBS = o, Responsavel = resp]
                            ),
                    validas = List.Select(tentativas, each [Data] <> null),
                    ordenadas =
                        List.Sort(
                            validas,
                            (a,b) =>
                                if a[Data] > b[Data] then -1
                                else if a[Data] < b[Data] then 1
                                else if a[Idx] > b[Idx] then -1
                                else if a[Idx] < b[Idx] then 1
                                else 0
                        ),
                    ultimo =
                        if List.Count(ordenadas) = 0
                        then [Data=null, OBS=null, Responsavel=null]
                        else [Data=ordenadas{0}[Data], OBS=ordenadas{0}[OBS], Responsavel=ordenadas{0}[Responsavel]]
                in
                    ultimo,
            type record
        ),

    ExpUltimoContato =
        Table.ExpandRecordColumn(
            AddUltimoContato,
            "UltimoContato",
            {"Data","OBS","Responsavel"},
            {"Data","OBS","Responsável"}
        ),

    // --------------------------------------------------
    // 6) QTD. TENTATIVAS E DIAS SEM CONTATO
    // --------------------------------------------------
    AddQtdTent =
        Table.AddColumn(
            ExpUltimoContato,
            "Qtd. tentativas",
            (r as record) as number =>
                if List.Count(IndicesValidos) = 0 then 0 else
                let
                    datas =
                        List.RemoveNulls(
                            List.Transform(
                                IndicesValidos,
                                (i) => try Date.From(Record.FieldOrDefault(r, "Data_" & Text.From(i), null)) otherwise null
                            )
                        )
                in
                    List.Count(datas),
            Int64.Type
        ),

    AddDiasSem =
        Table.AddColumn(
            AddQtdTent,
            "Dias sem contato",
            each if [Data] = null then null else Duration.Days(Date.From(DateTime.LocalNow()) - [Data]),
            Int64.Type
        ),

    // --------------------------------------------------
    // 7) SELEÇÃO FINAL (VISÃO ENXUTA)
    // --------------------------------------------------
    ColsFinaisDesejadas = {
        "ID","Parent task ID","Created on","Plano","Nome","Matrícula",
        "Data","OBS","Responsável",
        "Qtd. tentativas","Dias sem contato"
    },
    ColsExistentes = List.Intersect({ColsFinaisDesejadas, Table.ColumnNames(AddDiasSem)}),

    Saida    = Table.SelectColumns(AddDiasSem, ColsExistentes, MissingField.Ignore),
    Ordenado = Table.Sort(Saida, {{"ID", Order.Ascending}}),
    RemDup   = Table.Distinct(Ordenado, {"ID"})
in
    RemDup
