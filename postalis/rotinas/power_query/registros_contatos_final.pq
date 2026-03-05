// =====================================================================================
// CONSULTA: Registros_Contatos_Final
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Painel do dia (Registros_Contatos_Auto) + dados da Manual (1 por ID)
//             DIAS SEM CONTATO: pelo DIA DO PAINEL (max DataArquivo)
//             DATA ENTRADA: da Manual (real do ID)
// BASELINE OFICIAL: 01/02/2026 (nunca usar data do computador como fallback)
// VERSÃO: 03/03/2026 (DiaPainel sem Date/Now + clamp de dias)
// RESPONSÁVEL: Gabi (Postalis)
// =====================================================================================
let
    // 0) BASELINE (fixo)
    Baseline = #date(2026, 2, 1),

    // 1) DIA DO PAINEL (max DataArquivo na Correção_Automática; fallback = Baseline)
    DiaPainel =
        let
            t = Table.TransformColumnTypes(Correção_Automática, {{"DataArquivo", type date}}, "pt-BR"),
            ds0 = try List.RemoveNulls(Table.Column(t, "DataArquivo")) otherwise {},
            ds  = List.Select(ds0, each _ >= Baseline)
        in
            if List.Count(ds) = 0 then Baseline else List.Max(ds),

    // 2) FONTE (Auto)
    Fonte0 =
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

    // 2.1) Se a Fonte já tiver "Data Entrada", remove (para não virar Data Entrada2)
    Fonte =
        if List.Contains(Table.ColumnNames(Fonte0), "Data Entrada") then
            Table.RemoveColumns(Fonte0, {"Data Entrada"})
        else
            Fonte0,

    // 3) MANUAL 1 por ID
    Manual_Tipada = Registros_Contatos_Manual,
    Manual_Ordenada = try Table.Sort(Manual_Tipada, {{"Created on", Order.Descending}}) otherwise Manual_Tipada,
    Manual_1porID = Table.Distinct(Manual_Ordenada, {"ID"}),
    ManualNames = Table.ColumnNames(Manual_1porID),

    // 4) Índices 1..10 existentes
    IndicesFixos = {1..10},
    IndicesValidos = List.Select(IndicesFixos, (i) => List.Contains(ManualNames, "Data_" & Text.From(i))),

    ColsFixas = {"Observação Extra","E-mail","Telefone","Data Entrada"},
    ColsDinamicas =
        List.Combine(
            List.Transform(
                IndicesValidos,
                (i) => {"Data_" & Text.From(i),"Tipo_" & Text.From(i),"OBS_" & Text.From(i),"Responsável_" & Text.From(i)}
            )
        ),
    ManualToExpand = List.Intersect({List.Combine({ColsFixas, ColsDinamicas}), ManualNames}),

    // 5) MERGE Manual
    MergeManual =
        Table.NestedJoin(
            Fonte,
            {"ID"},
            Manual_1porID,
            {"ID"},
            "Manual",
            JoinKind.LeftOuter
        ),

    // 5.1) Expande a Manual, mas "Data Entrada" entra como "Data Entrada_Manual" (sem conflito)
    ManualExpandPairs =
        List.Transform(
            ManualToExpand,
            (c) => if c = "Data Entrada" then {"Data Entrada", "Data Entrada_Manual"} else {c, c}
        ),

    ExpManual =
        if List.Count(ManualToExpand) = 0 then
            MergeManual
        else
            Table.ExpandTableColumn(MergeManual, "Manual", ManualToExpand, List.Transform(ManualExpandPairs, each _{1})),

    // 6) Último contato
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
                                        [Idx=i, Data=d, OBS=o, Responsavel=resp]
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

    // 7) Data Entrada (da Manual) tipada e renomeada para "Data Entrada"
    AddDataEntradaDate =
        if List.Contains(Table.ColumnNames(ExpUltimoContato), "Data Entrada_Manual") then
            Table.RenameColumns(
                Table.TransformColumns(
                    ExpUltimoContato,
                    {{"Data Entrada_Manual", each try Date.From(_) otherwise null, type date}}
                ),
                {{"Data Entrada_Manual", "Data Entrada"}}
            )
        else
            ExpUltimoContato,

    // 8) Qtd tentativas e dias sem contato (pelo DiaPainel)
    AddQtdTent =
        Table.AddColumn(
            AddDataEntradaDate,
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
            each
                let
                    dUlt = [Data],
                    diff = if dUlt = null then null else Duration.Days(DiaPainel - dUlt)
                in
                    if diff = null then null else if diff < 0 then 0 else diff,
            Int64.Type
        ),

    // 9) Saída enxuta
    ColsFinaisDesejadas = {
        "ID","Parent task ID","Created on","Plano","Nome","Matrícula",
        "Data","OBS","Responsável",
        "Qtd. tentativas","Dias sem contato",
        "Data Entrada"
    },
    ColsExistentes = List.Intersect({ColsFinaisDesejadas, Table.ColumnNames(AddDiasSem)}),

    Saida    = Table.SelectColumns(AddDiasSem, ColsExistentes, MissingField.Ignore),
    Ordenado = Table.Sort(Saida, {{"ID", Order.Ascending}}),
    RemDup   = Table.Distinct(Ordenado, {"ID"})
in
    RemDup
