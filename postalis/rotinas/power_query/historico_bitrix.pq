// =====================================================================================
// CONSULTA: Histórico_Bitrix
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Montar histórico consolidado das extrações Bitrix encontradas na pasta,
//             preservando ID, Parent task ID, datas e arquivo de origem.
// ORIGEM DOS DADOS: Pasta de extrações Bitrix (arquivos GT_PP_ ou “Extração Bitrix”)
// DESTINO: Tabela de histórico de tarefas/contatos
// VERSÃO: Atualizado em 13/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Mantém histórico desde 02/02/2026.
// - Lê SOMENTE a pasta definida (Folder.Contents) — não varre subpastas.
// - Para cada dia (DataArquivo), mantém APENAS o arquivo mais recente (Date modified).
// - NÃO trava em "Sheet1": seleciona automaticamente a aba/tabela que contém a coluna ID.
// - Colunas extração (visíveis): ParteHora, HoraNum, MinutoNum, DataHoraExtracao.
// =====================================================================================
let
    ColsObrigatorias = {
        "ID","Parent task ID","Estágio","Tags","Status",
        "Created on","Modified on","Task","Description",
        "NomeArquivo","DataArquivo"
    },

    // ✅ PASTA CORRETA (onde estão os GT_PP_...)
    Pasta = "C:\\Users\\gabriela.martins\\OneDrive - Postalis\\Arquivos de Victor Six Madureira Guedes - Extracao GT Pensao e Peculio\\Stage - Aguardar documentacao",
    Fonte = Folder.Contents(Pasta),

    // 1) Só arquivos (sem pastas) + só Excel + sem ~$ 
    SoExcel =
        Table.SelectRows(
            Fonte,
            each
                (try [Attributes][Directory] otherwise false) = false
                and [Extension] <> null
                and List.Contains({".xlsx",".xls",".xlsm"}, Text.Lower([Extension]))
                and not Text.StartsWith([Name], "~$")
        ),

    // 2) Filtra arquivos de extração (GT_PP_ ou "Extração Bitrix")
    GTpp = Table.SelectRows(SoExcel, each Text.StartsWith([Name], "GT_PP_")),
    FonteArquivosBruta =
        if Table.RowCount(GTpp) = 0
        then Table.SelectRows(SoExcel, each Text.StartsWith([Name], "Extração Bitrix"))
        else GTpp,

    // 3) DataArquivo (dia) = Date modified
    AddDataArquivo =
        Table.AddColumn(
            FonteArquivosBruta,
            "DataArquivo",
            each try Date.From([#"Date modified"]) otherwise null,
            type date
        ),

    // 4) Desde 02/02/2026
    Desde0202 =
        Table.SelectRows(
            AddDataArquivo,
            each [DataArquivo] <> null and [DataArquivo] >= #date(2026, 2, 2)
        ),

    // 5) Um arquivo por dia (mais recente)
    OrdenaArquivos =
        Table.Sort(
            Desde0202,
            {{"DataArquivo", Order.Ascending}, {"Date modified", Order.Descending}, {"Name", Order.Descending}}
        ),

    UmPorDia =
        Table.Group(
            OrdenaArquivos,
            {"DataArquivo"},
            {{"Top1", each Table.FirstN(_, 1), type table}}
        ),

    FonteArquivos =
        Table.ExpandTableColumn(
            UmPorDia,
            "Top1",
            {"Content","Name","Date modified","Extension","Attributes","Folder Path"},
            {"Content","Name","Date modified","Extension","Attributes","Folder Path"}
        ),

    // 6) Abre workbook
    AbreWB = Table.AddColumn(FonteArquivos, "WB", each Excel.Workbook([Content], true)),
    ExpWB  = Table.ExpandTableColumn(AbreWB, "WB", {"Name","Data","Kind"}, {"SheetName","Data","Kind"}),
    SoPlan = Table.SelectRows(ExpWB, each [Kind] = "Sheet" or [Kind] = "Table"),

    // 7) Seleciona automaticamente a aba/tabela que tem a coluna ID
    TemID =
        Table.SelectRows(
            SoPlan,
            each
                let
                    cols = try Table.ColumnNames([Data]) otherwise {},
                    colsNorm = List.Transform(cols, each Text.Upper(Text.Trim(_)))
                in
                    List.Contains(colsNorm, "ID")
        ),

    _Valida =
        if Table.RowCount(TemID) = 0
        then error "Nenhuma aba/tabela com coluna ID foi encontrada nas extrações."
        else TemID,

    // 8) Expande os dados usando as colunas do primeiro item válido
    ExpDados = Table.ExpandTableColumn(_Valida, "Data", Table.ColumnNames(_Valida{0}[Data])),

    // 9) Normaliza nome da coluna ID se veio como "Id" ou "ID " etc.
    ColunasAtuais = Table.ColumnNames(ExpDados),
    ColunasNorm = List.Transform(ColunasAtuais, each Text.Upper(Text.Trim(_))),
    PosID = List.PositionOf(ColunasNorm, "ID"),
    Base =
        if PosID >= 0 and ColunasAtuais{PosID} <> "ID"
        then Table.RenameColumns(ExpDados, {{ColunasAtuais{PosID}, "ID"}}, MissingField.Ignore)
        else ExpDados,

    // 10) Padroniza Estágio
    StdEstagio1 =
        if Table.HasColumns(Base, {"Stage"})
        then Table.RenameColumns(Base, {{"Stage","Estágio"}}, MissingField.Ignore)
        else Base,

    StdEstagio2 =
        if Table.HasColumns(StdEstagio1, {"Estágio"})
        then StdEstagio1
        else Table.AddColumn(StdEstagio1, "Estágio", each null, type text),

    // 11) NomeArquivo (origem)
    AddNome = Table.AddColumn(StdEstagio2, "NomeArquivo", each [Name], type text),

    // 12) Remove meta desnecessária
    MetaCols = {"Content","Extension","Date modified","Attributes","Folder Path","SheetName","Kind","Name"},
    SemMeta =
        Table.RemoveColumns(
            AddNome,
            List.Intersect({MetaCols, Table.ColumnNames(AddNome)})
        ),

    // 13) Garante Parent task ID
    GarantirParent =
        if Table.HasColumns(SemMeta, {"Parent task ID"})
        then SemMeta
        else Table.AddColumn(SemMeta, "Parent task ID", each null, type text),

    // 14) Tipagem + seleção final (contrato)
    Desejados = {
        {"ID", type text},
        {"Parent task ID", type text},
        {"Task", type text},
        {"Description", type text},
        {"Status", type text},
        {"Tags", type text},
        {"Created on", type datetime},
        {"Modified on", type datetime},
        {"Estágio", type text},
        {"NomeArquivo", type text},
        {"DataArquivo", type date}
    },

    NomesAtuais2 = Table.ColumnNames(GarantirParent),
    PairsExist  = List.Select(Desejados, each List.Contains(NomesAtuais2, _{0})),
    Tipado      = Table.TransformColumnTypes(GarantirParent, PairsExist, "pt-BR"),

    ColsExistentes = List.Intersect({ColsObrigatorias, Table.ColumnNames(Tipado)}),
    Saida          = Table.SelectColumns(Tipado, ColsExistentes, MissingField.Ignore),

    // 15) Deduplica por (ID, DataArquivo)
    LinhasClassificadas = Table.Sort(Saida, {{"NomeArquivo", Order.Descending}}),
    DuplicatasRemovidas = Table.Distinct(LinhasClassificadas, {"ID", "DataArquivo"}),

    // 16) Extrair ParteHora (último token do NomeArquivo antes de .xlsx)
    AddParteHora =
        Table.AddColumn(
            DuplicatasRemovidas,
            "ParteHora",
            each
                let
                    base = try Text.BeforeDelimiter([NomeArquivo], ".xlsx") otherwise null,
                    tok  = if base <> null then try List.Last(Text.Split(base, "_")) otherwise null else null
                in
                    tok,
            type text
        ),

    // 17) HoraNum / MinutoNum (a partir de ParteHora, ex: 06h16min)
    AddHoraNum =
        Table.AddColumn(
            AddParteHora,
            "HoraNum",
            each try Number.FromText(Text.BeforeDelimiter([ParteHora], "h")) otherwise null,
            Int64.Type
        ),

    AddMinutoNum =
        Table.AddColumn(
            AddHoraNum,
            "MinutoNum",
            each try Number.FromText(Text.BeforeDelimiter(Text.AfterDelimiter([ParteHora], "h"), "min")) otherwise null,
            Int64.Type
        ),

    // 18) DataHoraExtracao = DataArquivo + Hora/Minuto (DateTime real)
    AddDataHoraExtracao =
        Table.AddColumn(
            AddMinutoNum,
            "DataHoraExtracao",
            each
                try
                    if [DataArquivo] <> null and [HoraNum] <> null and [MinutoNum] <> null
                    then #datetime(Date.Year([DataArquivo]), Date.Month([DataArquivo]), Date.Day([DataArquivo]), [HoraNum], [MinutoNum], 0)
                    else null
                otherwise null,
            type datetime
        ),

    // 19) Reordenação final (mantém as colunas novas VISÍVEIS)
    ColunasFinal =
        Table.ReorderColumns(
            AddDataHoraExtracao,
            {"ID","Parent task ID","Estágio","Tags","Status","Created on","Modified on","Task","Description","NomeArquivo","DataArquivo","ParteHora","HoraNum","MinutoNum","DataHoraExtracao"},
            MissingField.Ignore
        )
in
    ColunasFinal
