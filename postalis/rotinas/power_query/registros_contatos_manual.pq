// =====================================================================================
// CONSULTA: Registros_Contatos_Manual
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: Expor a tabela de preenchimento manual de contatos com tipagem correta.
// ORIGEM DOS DADOS: Tabela Excel "Registros_Contatos_Manual" (ThisWorkbook)
// DESTINO: Consulta auxiliar usada em Registros_Contatos_Final e VBA
// VERSÃO: Atualizado em 09/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - NÃO altera dados: apenas lê a tabela já existente no Excel.
// – Datas 1 a 5 (date), Created on (datetime), blindagem completa
// =====================================================================================
let
    Fonte = Excel.CurrentWorkbook(){[Name = "Registros_Contatos_Manual"]}[Content],

    TiposDesejados = {
        {"ID", type text},
        {"Parent task ID", type text},
        {"Created on", type datetime},
        {"Matrícula", type text},
        {"Nome", type text},
        {"Observação Extra", type text},
        {"E-mail", type text},
        {"Telefone", type text},

        {"Data_1", type date},
        {"OBS_1", type text},
        {"Responsável_1", type text},

        {"Data_2", type date},
        {"OBS_2", type text},
        {"Responsável_2", type text},

        {"Data_3", type date},
        {"OBS_3", type text},
        {"Responsável_3", type text},

        {"Data_4", type date},
        {"OBS_4", type text},
        {"Responsável_4", type text},

        {"Data_5", type date},
        {"OBS_5", type text},
        {"Responsável_5", type text}
    },

    ColunasExistentes = Table.ColumnNames(Fonte),
    TiposAplicaveis = List.Select(TiposDesejados, each List.Contains(ColunasExistentes, _{0})),

    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte, TiposAplicaveis, "pt-BR")
in
    #"Tipo Alterado"
