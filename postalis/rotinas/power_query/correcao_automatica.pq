// =====================================================================================
// CONSULTA: Correção_Automática
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Produzir a base "limpa" da extração Bitrix, com as colunas padronizadas usadas no fluxo.
// ORIGEM DOS DADOS: Bruta_Bitrix
// DESTINO: Tabela Correção_Automática no Excel
// VERSÃO: v1.0 – atualizado em 05/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Garante a existência de Parent task ID.
// - Mantém apenas as colunas necessárias para o fluxo de contatos.
// =====================================================================================
let
    Fonte0=Bruta_Bitrix,
    GarantirParent=if Table.HasColumns(Fonte0,{"Parent task ID"}) then Fonte0 else Table.AddColumn(Fonte0,"Parent task ID",each null,type text),
    ColunasNecessarias=Table.SelectColumns(GarantirParent,{"ID","Parent task ID","Estágio","Tags","Status","Created on","Modified on","Task","Description","NomeArquivo","DataArquivo"},MissingField.Ignore),
    Tipos=Table.TransformColumnTypes(ColunasNecessarias,{{"Created on",type datetime},{"Modified on",type datetime},{"DataArquivo",type date}},"pt-BR")
in
    Tipos
