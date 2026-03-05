// =====================================================================================
// CONSULTA: Registros_Contatos_Manual
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Expor a tabela de preenchimento manual de contatos com tipagem correta.
// ORIGEM DOS DADOS: Tabela Excel "Registros_Contatos_Manual" (ThisWorkbook)
// DESTINO: Consulta auxiliar usada em Registros_Contatos_Final e VBA
// VERSÃO: Atualizado em 02/03/2026 (Caminho A - 1..10 + colunas novas)
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - NÃO altera dados: apenas lê a tabela já existente no Excel.
// - Blindagem completa: só aplica tipagem se a coluna existir.
// =====================================================================================
let
    Fonte = Excel.CurrentWorkbook(){[Name = "Registros_Contatos_Manual"]}[Content],

    TiposDesejados = {
        {"ID", type text},
        {"Parent task ID", type text},
        {"Created on", type datetime},
        {"Matrícula", type text},
        {"Nome", type text},

        {"Data Entrada", type date},
        {"Data Saída", type date},
        {"Data Retorno", type date},
        {"Status", type text},
        {"Tentativas_Realizadas", Int64.Type},

        {"Observação Extra", type text},
        {"E-mail", type text},
        {"Telefone", type text},
        {"Contato_Extra", type text},

        {"Data_1", type date}, {"Tipo_1", type text}, {"OBS_1", type text}, {"Responsável_1", type text},
        {"Data_2", type date}, {"Tipo_2", type text}, {"OBS_2", type text}, {"Responsável_2", type text},
        {"Data_3", type date}, {"Tipo_3", type text}, {"OBS_3", type text}, {"Responsável_3", type text},
        {"Data_4", type date}, {"Tipo_4", type text}, {"OBS_4", type text}, {"Responsável_4", type text},
        {"Data_5", type date}, {"Tipo_5", type text}, {"OBS_5", type text}, {"Responsável_5", type text},
        {"Data_6", type date}, {"Tipo_6", type text}, {"OBS_6", type text}, {"Responsável_6", type text},
        {"Data_7", type date}, {"Tipo_7", type text}, {"OBS_7", type text}, {"Responsável_7", type text},
        {"Data_8", type date}, {"Tipo_8", type text}, {"OBS_8", type text}, {"Responsável_8", type text},
        {"Data_9", type date}, {"Tipo_9", type text}, {"OBS_9", type text}, {"Responsável_9", type text},
        {"Data_10", type date}, {"Tipo_10", type text}, {"OBS_10", type text}, {"Responsável_10", type text}
    },

    ColunasExistentes = Table.ColumnNames(Fonte),
    TiposAplicaveis = List.Select(TiposDesejados, each List.Contains(ColunasExistentes, _{0})),

    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte, TiposAplicaveis, "pt-BR")
in
    #"Tipo Alterado"
