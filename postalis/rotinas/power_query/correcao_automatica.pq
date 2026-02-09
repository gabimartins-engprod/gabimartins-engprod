// =====================================================================================
// CONSULTA: Correção_Automática
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Produzir a base "limpa" da extração Bitrix, com as colunas padronizadas
//             usadas no restante do fluxo (Registros_Contatos_Auto, etc.).
// ORIGEM DOS DADOS: Consulta Bruta_Bitrix
// DESTINO: Tabela Correção_Automática no Excel
// VERSÃO: v1.1 – atualizado em 09/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Valida a existência da coluna essencial ID.
// - Garante a existência da coluna Parent task ID.
// - Mantém apenas as colunas necessárias para o fluxo.
// - Força consistência de tipos (ID e Parent task ID como texto).
// - Converte DataArquivo para data de forma robusta.
// =====================================================================================
let
    Fonte0 = Bruta_Bitrix,

    ValidarID =
        if Table.HasColumns(Fonte0, {"ID"})
        then Fonte0
        else error "Correção_Automática: coluna essencial 'ID' não encontrada na Bruta_Bitrix.",

    GarantirParent =
        if Table.HasColumns(ValidarID, {"Parent task ID"})
        then ValidarID
        else Table.AddColumn(ValidarID, "Parent task ID", each null, type text),

    ColunasNecessarias =
        Table.SelectColumns(
            GarantirParent,
            {
                "ID","Parent task ID","Estágio","Tags","Status",
                "Created on","Modified on","Task","Description",
                "NomeArquivo","DataArquivo"
            },
            MissingField.Ignore
        ),

    TiposBase =
        Table.TransformColumnTypes(
            ColunasNecessarias,
            {
                {"ID", type text},
                {"Parent task ID", type text},
                {"Created on", type datetime},
                {"Modified on", type datetime}
            },
            "pt-BR"
        ),

    DataArquivoParaDate =
        if Table.HasColumns(TiposBase, {"DataArquivo"})
        then
            Table.TransformColumns(
                TiposBase,
                {
                    {
                        "DataArquivo",
                        each
                            let v = _
                            in
                                if v = null then null
                                else if Value.Is(v, type date) then v
                                else if Value.Is(v, type datetime) then DateTime.Date(v)
                                else try Date.From(v) otherwise null,
                        type date
                    }
                }
            )
        else
            TiposBase
in
    DataArquivoParaDate
