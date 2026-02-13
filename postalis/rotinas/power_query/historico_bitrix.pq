// =====================================================================================
// CONSULTA: Histórico_Bitrix
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Montar histórico consolidado das extrações Bitrix encontradas na pasta,
//             preservando ID, Parent task ID, datas e arquivo de origem.
// ORIGEM DOS DADOS: Pasta de extrações Bitrix (arquivos GT_PP_ ou “Extração Bitrix”)
// DESTINO: Tabela de histórico de tarefas/contatos (tblHistoricoBitrix)
// VERSÃO: Atualizado em 12/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Mantém histórico desde 02/02/2026.
// - Lê SOMENTE a pasta definida (Folder.Contents) — não varre subpastas.
// - Regra oficial: para cada dia (DataArquivo), mantém APENAS a ÚLTIMA extração do dia,
//   definida pela coluna DataHoraExtracao (DateTime real derivado do NomeArquivo).
// - NÃO trava em "Sheet1": seleciona automaticamente a aba/tabela que contém a coluna ID.
// =====================================================================================
let
    ColsObrigatorias = {
        "ID","Parent task ID","Estágio","Tags","Status",
        "Created on","Modified on","Task","Description",
        "NomeArquivo","DataArquivo","DataHoraExtracao"
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

    // 3) DataArquivo (dia) = Date modified (data do dia) — mantém coerência com a pasta
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

    // 5) Derivar DataHoraExtracao (DateTime real) a partir do NomeArquivo: ..._06h16min.xlsx
    AddParteHora =
        Table.AddColumn(
            Desde0202,
            "ParteHora",
            each
                let
                    base = Text.BeforeDelimiter([Name], ".xlsx"),
                    last = List.Last(Text.Split(base, "_"))
                in
                    last,
            type text
        ),

    AddHoraNum =
        Table.AddColumn(
            AddParteHora,
            "HoraNum",
            each
                let
                    s = [ParteHora],
                    hTxt = Text.BeforeDelimiter(s, "h"),
                    hNum = try Number.FromText(hTxt) otherwise null
                in
                    hNum,
            Int64.Type
        ),

    AddMinutoNum =
        Table.AddColumn(
            AddHoraNum,
            "MinutoNum",
            each
                let
                    s = [ParteHora],
                    mTxt = Text.BetweenDelimiters(s, "h", "min"),
                    mNum = try Number.FromText(mTxt) otherwise null
                in
                    mNum,
            Int64.Type
        ),

    AddDataHoraExtracao =
        Table.AddColumn(
            AddMinutoNum,
            "DataHoraExtracao",
            each
                let
                    d = [DataArquivo],
                    h = [HoraNum],
                    m = [MinutoNum]
                in
                    if d <> null and h <> null and m <> null
                    then #datetime(Date.Year(d), Date.Month(d), Date.Day(d), h, m, 0)
                    else null,
            type datetime
        ),

    // 6) Um arquivo por dia (ÚLTIMA extração do dia = maior DataHoraExtracao)
    OrdenaArquivos =
        Table.Sort(
            AddDataHoraExtracao,
            {{"DataArquivo", Order.Ascending}, {"DataHoraExtracao", Order.Descending}, {"Name", Order.Descending}}
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
            {"Content","Name","Date modified","Extension","Attributes","Folder Path","DataHoraExtracao"},
            {"Content","Name","Date modified","Extension","Attributes","Folder Path","DataHoraExtracao"}
        ),

    // 7) Abre workbook
    AbreWB = Table.AddColumn(FonteArquivos, "WB", each Excel.Workbook([Content], true)),
    ExpWB  = Table.ExpandTableColumn(AbreWB, "WB", {"Name","Data","Kind"}, {"SheetName","Data","Kind"}),
    SoPlan = Table.SelectRows(ExpWB, each [Kind] = "Sheet" or [Kind] = "Table"),

    // 8) Seleciona automaticamente a aba/tabela que tem a coluna ID
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

    // 9) Expande os dados usando as colunas do primeiro item válido
    ExpDados = Table.ExpandTableColumn(_Valida, "Data", Table.ColumnNames(_Valida{0}[Data])),

    // 10) Normaliza nome da coluna ID se veio como "Id" ou "ID " etc.
    ColunasAtuais = Table.ColumnNames(ExpDados),
    ColunasNorm = List.Transform(ColunasAtuais, each Text.Upper(Text.Trim(_))),
    PosID = List.PositionOf(ColunasNorm, "ID"),
    Base =
        if PosID >= 0 and ColunasAtuais{PosID} <> "ID"
        then Table.RenameColumns(ExpDados, {{ColunasAtuais{PosID}, "ID"}}, MissingField.Ignore)
        else ExpDados,

    // 11) Padroniza Estágio
    StdEstagio1 =
        if Table.HasColumns(Base, {"Stage"})
        then Table.RenameColumns(Base, {{"Stage","Estágio"}}, MissingField.Ignore)
        else Base,

    StdEstagio2 =
        if Table.HasColumns(StdEstagio1, {"Estágio"})
        then StdEstagio1
        else Table.AddColumn(StdEstagio1, "Estágio", each null, type text),

    // 12) NomeArquivo (origem)
    AddNome = Table.AddColumn(StdEstagio2, "NomeArquivo", each [Name], type text),

    // 13) Remove meta desnecessária
    MetaCols = {"Content","Extension","Date modified","Attributes","Folder Path","SheetName","Kind","Name","WB"},
    SemMeta =
        Table.RemoveColumns(
            AddNome,
            List.Intersect({MetaCols, Table.ColumnNames(AddNome)})
        ),

    // 14) Garante Parent task ID
    GarantirParent =
        if Table.HasColumns(SemMeta, {"Parent task ID"})
        then SemMeta
        else Table.AddColumn(SemMeta, "Parent task ID", each null, type text),

    // 15) Tipagem + seleção final
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
        {"DataArquivo", type date},
        {"DataHoraExtracao", type datetime}
    },

    NomesAtuais2 = Table.ColumnNames(GarantirParent),
    PairsExist  = List.Select(Desejados, each List.Contains(NomesAtuais2, _{0})),
    Tipado      = Table.TransformColumnTypes(GarantirParent, PairsExist, "pt-BR"),

    ColsExistentes = List.Intersect({ColsObrigatorias, Table.ColumnNames(Tipado)}),
    Saida          = Table.SelectColumns(Tipado, ColsExistentes, MissingField.Ignore),

    // 16) Deduplica por (ID, DataArquivo) — mantém 1 linha por ID por dia (já na ÚLTIMA extração do dia)
    LinhasClassificadas = Table.Sort(Saida, {{"NomeArquivo", Order.Descending}}),
    DuplicatasRemovidas = Table.Distinct(LinhasClassificadas, {"ID", "DataArquivo"})
in
    DuplicatasRemovidas
