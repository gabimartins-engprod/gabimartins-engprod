// =====================================================================================
// CONSULTA: Bruta_Bitrix
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE:
//   Localizar automaticamente o arquivo de extração Bitrix mais recente em pasta definida,
//   identificar a aba ou tabela correta, padronizar colunas principais e garantir a
//   existência de todas as colunas obrigatórias para uso nas demais consultas.
//
// ORIGEM DOS DADOS:
//   Pasta local contendo arquivos Excel de extração do Bitrix.
//   Prioridade para arquivos iniciados com "GT_PP_".
//   Alternativa para arquivos iniciados com "Extração Bitrix".
//
// DESTINO:
//   Consulta de staging (apenas conexão), utilizada como base por:
//     - Correção_Automática
//     - Registros_Contatos_Auto
//     - Histórico_Bitrix
//
// COMPORTAMENTO TÉCNICO:
//   - Seleciona o arquivo mais recente por "Date modified"
//   - Normaliza variações de cabeçalhos (ID, Estágio, Status, etc.)
//   - Garante colunas obrigatórias, criando-as como nulas quando ausentes
//   - Adiciona metadados: NomeArquivo e DataArquivo
//
// VERSÃO: v1.1 – atualizado em 09/02/2026
// RESPONSÁVEL: Gabi (Postalis)
//
// OBSERVAÇÃO:
//   O caminho da pasta de extração encontra-se definido diretamente no código nesta versão.
//   Recomenda-se a parametrização para execução em outros ambientes.
// =====================================================================================
