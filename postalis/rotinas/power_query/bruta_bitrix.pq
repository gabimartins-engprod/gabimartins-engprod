// =====================================================================================
// CONSULTA: Bruta_Bitrix
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Ler a extração mais recente do Bitrix na pasta especificada, normalizar
//             cabeçalhos e garantir as colunas padrão usadas nas demais consultas.
// ORIGEM DOS DADOS: Pasta com arquivos Excel de extração Bitrix (GT_PP_ ou "Extração Bitrix")
// DESTINO: Tabela padronizada para uso em Correção_Automática e Histórico_Bitrix
// VERSÃO: v1.0 – atualizado em 05/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Seleciona o arquivo mais recente na pasta.
// - Normaliza nomes de colunas (Id→ID, Stage→Estágio, etc.).
// - Garante colunas obrigatórias mesmo que faltem na origem.
// =====================================================================================
let
    ColsObrigatorias={"ID","Parent task ID","Estágio","Tags","Status","Created on","Modified on","Task","Description"},
    ModoEstrito=false,
    Pasta="C:\\CAMINHO_DA_PASTA\\Stage - Aguardar documentacao",
    Fonte=Folder.Files(Pasta),
    SoExcel=Table.SelectRows(Fonte,each let ext=Text.Lower([Extension]) in (ext=".xlsx" or ext=".xls" or ext=".xlsm") and not Text.StartsWith([Name],"~$")),
    GTpp=Table.SelectRows(SoExcel,each Text.StartsWith([Name],"GT_PP_")),
    FonteArquivos=if Table.RowCount(GTpp)=0 then Table.SelectRows(SoExcel,each Text.StartsWith([Name],"Extração Bitrix")) else GTpp,
    Ordena=Table.Sort(FonteArquivos,{{"Date modified",Order.Descending}}),
    Ultimo=if Table.RowCount(Ordena)=0 then error "Nenhum arquivo de extração encontrado na pasta definida." else Table.FirstN(Ordena,1),
    NomeArquivo=Ultimo{0}[Name],
    DataArquivoDT=Ultimo{0}[Date modified],
    AbreWB=Table.AddColumn(Ultimo,"WB",each Excel.Workbook([Content],true)),
    ExpWB=Table.ExpandTableColumn(AbreWB,"WB",{"Name","Data","Kind"},{"SheetName","Data","Kind"}),
    SoPlan=Table.SelectRows(ExpWB,each [Kind]="Sheet" or [Kind]="Table"),
    ComID=Table.SelectRows(SoPlan,each Table.HasColumns([Data],{"ID"}) or Table.HasColumns([Data],{"Id"})),
    FonteEscolhida=if Table.RowCount(ComID)>0 then Table.FirstN(ComID,1) else let AddColCount=Table.AddColumn(SoPlan,"ColCount",each Table.ColumnCount([Data]),Int64.Type),OrdenaCols=Table.Sort(AddColCount,{{"ColCount",Order.Descending}}) in Table.FirstN(OrdenaCols,1),
    ColsOrigem=Table.ColumnNames(FonteEscolhida{0}[Data]),
    BaseExp=Table.ExpandTableColumn(FonteEscolhida,"Data",ColsOrigem),
    MapPairs={{"Id","ID"},{"id","ID"},{"Stage","Estágio"},{"stage","Estágio"},{"status","Status"},{"tags","Tags"},{"task","Task"},{"description","Description"},{"created on","Created on"},{"modified on","Modified on"},{"parent task id","Parent task ID"},{"estágio","Estágio"}},
    PairsQueExistem=List.Select(MapPairs,(p)=>Table.HasColumns(BaseExp,{p{0}}) and not Table.HasColumns(BaseExp,{p{1}})),
    Normalizado=if List.Count(PairsQueExistem)>0 then Table.RenameColumns(BaseExp,PairsQueExistem,MissingField.Ignore) else BaseExp,
    BaseID=if Table.HasColumns(Normalizado,{"Id"}) then Table.RenameColumns(Normalizado,{{"Id","ID"}}) else Normalizado,
    StdEstagio=if Table.HasColumns(BaseID,{"Stage"}) then Table.RenameColumns(BaseID,{{"Stage","Estágio"}}) else if Table.HasColumns(BaseID,{"Estágio"}) then BaseID else Table.AddColumn(BaseID,"Estágio",each null,type text),
    ColsAtuais=Table.ColumnNames(StdEstagio),
    Faltando=List.Difference(ColsObrigatorias,ColsAtuais),
    DepoisDeGarantir=if (ModoEstrito and List.Count(Faltando)>0) then error "Faltam colunas obrigatórias na extração: "&Text.Combine(Faltando,", ") else List.Accumulate(Faltando,StdEstagio,(state,colName)=>Table.AddColumn(state,colName,each null,type any)),
    Tipado=Table.TransformColumnTypes(DepoisDeGarantir,{{"ID",type text},{"Parent task ID",type text},{"Task",type text},{"Description",type text},{"Status",type text},{"Tags",type text},{"Created on",type datetime},{"Modified on",type datetime},{"Estágio",type text}},"pt-BR"),
    ComNome=Table.AddColumn(Tipado,"NomeArquivo",each NomeArquivo,type text),
    ComData=Table.AddColumn(ComNome,"DataArquivo",each Date.From(DataArquivoDT),type date),
    OrdemDesejada=List.Combine({ColsObrigatorias,{"NomeArquivo","DataArquivo"}}),
    ColsExistentes=List.Intersect({OrdemDesejada,Table.ColumnNames(ComData)}),
    Saida=Table.SelectColumns(ComData,ColsExistentes,MissingField.Ignore)
in
    Saida
