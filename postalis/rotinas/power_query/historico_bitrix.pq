// =====================================================================================
// CONSULTA: Histórico_Bitrix
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE:
//   Consolidar o histórico diário das extrações Bitrix, garantindo correspondência
//   exata entre cada arquivo de extração e o respectivo dia.
//
// ORIGEM DOS DADOS:
//   Pasta local contendo arquivos Excel de extração do Bitrix.
//
// DESTINO:
//   Tabela carregada em planilha (tblHistoricoBitrix), com aba mantida oculta.
//
// ATUALIZADO EM: 09/02/2026
// RESPONSÁVEL: Gabi (Postalis)
//
// LÓGICA PRINCIPAL:
//   - Lê apenas arquivos Excel válidos (.xlsx, .xls, .xlsm)
//   - Prioriza arquivos iniciados com "GT_PP_"
//   - Seleciona automaticamente um único arquivo por dia
//     (definido pelo campo Date modified)
//   - Identifica dinamicamente a aba ou tabela que contém a coluna "ID"
//   - Normaliza e padroniza colunas essenciais
//   - Adiciona NomeArquivo e DataArquivo
//   - Consolida histórico diário por ID (1 linha por ID por dia)
//
// OBSERVAÇÕES TÉCNICAS:
//   - Caso alguma coluna obrigatória não exista na origem, é criada com valor nulo
//   - A consulta não depende de nomes fixos de abas (ex.: Sheet1)
//   - Utilizada como base para:
//       * Registros_Contatos_Auto
//       * Registros_Contatos_Final
// =====================================================================================
