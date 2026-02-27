# Estrutura de Rotina — VBA (Rotina_Dados – Postalis)

Documento criado para organizar a lógica geral do VBA da rotina administrativa Rotina_Dados – Postalis, descrevendo responsabilidades, dependências e sequência de execução, sem uso de dados reais.

## Objetivo

Documentar a arquitetura do fluxo VBA, deixando claro:

quais módulos existem;

qual a responsabilidade de cada um;

como eles se conectam ao Power Query e às tabelas;

como o botão único executa a rotina ponta a ponta.

## Etapas gerais (fluxo VBA)

Abertura do arquivo

Preparação do Power Query (sem background)

Atualização das consultas (RefreshAll)

Aplicação de fórmulas e layout

Sincronização da base final (Dentro/Fora + datas)

Sincronização e manutenção da base manual

Geração/atualização do Log_Sincronizacao

Recalcular e atualizar pivôs

Finalização

## Observações

Documento de caráter organizacional e técnico.

A rotina principal é manual via botão AtualizarTudo_E_Sincronizar (módulo modExecucao).

A lógica diária de comparação é baseada em DataArquivo (date) como “dia oficial” e DataHoraExtracao (datetime) apenas para selecionar a última extração do dia, quando disponível.

 ---

### ThisWorkbook (EstaPastaDeTrabalho) — VERSÃO: 13/02/2026

Módulo de evento do Excel, responsável por preparar o ambiente quando o arquivo é aberto.

#### Finalidade

Garantir previsibilidade na execução do fluxo VBA, chamando a rotina que desativa execução em background das consultas do Power Query.

#### Responsabilidades principais

Executar automaticamente ao abrir o arquivo (Workbook_Open)

Chamar PQ_BackgroundOff para desligar refresh em background

#### Dependências

modPQPowerQuery.PQ_BackgroundOff

#### Tipo de execução

Automática (evento do Excel)

#### Observações técnicas

Usa On Error Resume Next para não impedir a abertura do arquivo caso haja alguma divergência/erro.

 ---

# modPQPowerQuery — VERSÃO: 23/02/2026

Módulo responsável por forçar o Power Query a não rodar em background, permitindo que o VBA controle o encadeamento de atualizações.

#### Finalidade

Evitar execução assíncrona do Power Query, garantindo que o VBA consiga fazer:

RefreshAll

CalculateUntilAsyncQueriesDone
com consistência.

#### Responsabilidades principais

Percorrer todas as abas e ListObjects

Se a tabela tiver QueryTable, setar:

BackgroundQuery = False

RefreshStyle = xlInsertDeleteCells

Percorrer ThisWorkbook.Connections e, quando aplicável:

OLEDBConnection.BackgroundQuery = False

ODBCConnection.BackgroundQuery = False

#### Origem/entrada

Estruturas do Excel (ListObjects, QueryTables e Connections)

#### Tipo de execução

Preparação técnica (configuração do ambiente)

#### Observações técnicas

Tratamento de erro com Debug.Print para diagnóstico sem travar a rotina.

 ---

# modLayoutPadrao — VERSÃO: 23/02/2026

Módulo de padronização visual e performance do layout, com AutoFit leve.

#### Finalidade

Aplicar layout padronizado nas abas principais sem o custo do AutoFit completo na planilha inteira.

Abas alvo

Extração Bitrix

Registros Contatos

Registros Contatos Manual

#### Responsabilidades principais

Ajustar altura de linhas:

Linha 1: altura maior (cabeçalho)

Linha 2: altura intermediária

Linhas 3 até o final: altura padrão

Aplicar AutoFit apenas nas colunas das tabelas, usando amostra de linhas

#### Controle de performance

AutoFit roda uma vez por dia, salvo execução forçada

Usa o Name do workbook:

RD_LastAutoFitDate

#### Tipo de execução

Ajuste visual pós-processamento (layout)

#### Observações técnicas

AutoFitListObject_Sample: pega Header + N primeiras linhas (amostragem) e faz AutoFit apenas nessas colunas.

 ---

# modAplicarFormulasPendencias — VERSÃO: 25/02/2026

Módulo responsável por aplicar fórmulas (PROCX/LET) na tabela Correção_Automática, espelhando informações preenchidas manualmente na tabela final.

#### Finalidade

Garantir que a base Correção_Automática carregue, via fórmula, os campos manuais existentes na Registros_Contatos_Final, centralizando o espelhamento.

#### Tabelas envolvidas

Destino: Correção_Automática

Fonte: Registros_Contatos_Final

#### Responsabilidades principais

Definir e manter mapa de colunas (COLS_MAP)

Criar colunas ausentes no destino quando necessário

Aplicar fórmulas por coluna:

campos de data com conversão robusta (DATA.VALOR quando texto)

campos de texto com retorno vazio quando nulo/zero

Aplicar formatos:

datas (ex.: dd/mmm/aa)

e-mail e telefone como texto

#### Otimização (TURBO)

Só reescreve a fórmula quando:

a coluna é nova, ou

a fórmula mudou (comparação normalizada)

#### Integração com layout

Wrapper Rotina_Sincronizar_E_Recalcular chama:

AplicarFormulasPendencias

modLayoutPadrao.AplicarLayoutPadraoTabelas

#### Observação (Opção A)

Este módulo não orquestra sincronizações (para evitar duplicidade)

A orquestração fica centralizada no modExecucao

---

# modSincronizarContatos — VERSÃO: 25/02/2026

Módulo central de sincronização de registros, responsável por manter consistência entre a base automática, a base final e a base manual.

#### Finalidade

Sincronizar Registros_Contatos_Final com Correção_Automática

Manter e alimentar Registros_Contatos_Manual com IDs e datas operacionais

#### Tabelas envolvidas

Correção_Automática (base automática tratada)

Registros_Contatos_Final (base oficial consolidada)

Registros_Contatos_Manual (base persistente manual)

Histórico: tblHistoricoBitrix (aba Histórico Bitrix)

#### Conceitos e regras

BASELINE conceitual: 01/02/2026

Backfill e ciclo percorrem somente dias com extração (baseados em DataArquivo)

DataHoraExtracao é opcional e, quando existe, serve para selecionar “última foto do dia”

#### Responsabilidades principais (por rotina)

(A) Sincronizar_AteX_RegistrosContatos

Garante colunas: Status, Data Entrada, Data Saída

Define status:

Dentro da extração quando ID está no conjunto atual

Fora da extração quando não está

Preenche:

Data Entrada quando vazio (dataRef)

Data Saída quando passa a “fora”

(B) Manual_IncluirNovosIDs_PeloFinal

Inclui IDs que existem no Final mas não existem na Manual

Copia campos chave: Parent task ID, Created on, Matrícula, Nome

(C) Manual_IncluirIDs_PeloHistorico_0202

Inclui IDs novos diretamente do Histórico desde o baseline

Critério: pega a linha mais recente por ID (maior DataArquivo; desempate por DataHoraExtracao)

(D) Manual_Backfill_Pendentes_0202 (Backfill leve)

Preenche pendências de datas na Manual:

Data Entrada

Data Saída

Data Retorno

Limpa datas inválidas (dia sem extração)

Calcula saída/retorno percorrendo apenas extrDates (dias com extração)

#### Tipo de execução

Execução por etapas, chamada pelo modExecucao

#### Observações técnicas

IDs são tratados como texto (NumberFormat = "@")

Usa Scripting.Dictionary para performance (sets e presença diária)

Backfill é limitado a pendentes (mais leve e seguro).

 ---

# modExecucao — VERSÃO: 25/02/2026

Módulo orquestrador do sistema (botão único), responsável por rodar a rotina completa com proteção anti-duplicidade.

#### Finalidade

Executar o fluxo ponta a ponta: atualizar consultas, aplicar regras, sincronizar tabelas e gerar log histórico.

#### Botão único

AtualizarTudo_E_Sincronizar

#### Etapas executadas (sequência)

PQ_BackgroundOff (opcional)

RefreshAllAndWait (Power Query)

modAplicarFormulasPendencias.Rotina_Sincronizar_E_Recalcular

modSincronizarContatos.Sincronizar_AteX_RegistrosContatos

Manual_IncluirNovosIDs_PeloFinal + Manual_IncluirIDs_PeloHistorico_0202

Manual_Backfill_Rapido_0202

BuildLog_FromHistorico_FillMissingDays (desde 01/02/2026)

Recalcular + FixTotalDia_IfExists

Atualizar tabelas dinâmicas (RefreshAllPivots)

#### Regras críticas do Log (Histórico)

Referência de “dia”: DataArquivo

Seleção da “foto válida do dia”:

se existe DataHoraExtracao, usa apenas a maior daquele dia

se não existe, usa todas as linhas do dia por DataArquivo

Gera blocos por dia:

INICIO

(RETORNOU / SAIU / ENTROU) quando aplicável

RESUMO

FIM

Para dias sem extração:

registra RESUMO com alerta para verificação

#### Anti-duplicidade por clique (Run once por etapa)

Cache em dictionary gRanSteps

Cada etapa roda no máximo uma vez por execução do botão

#### Correção TotalDia

FixTotalDia_IfExists preenche TotalDia somente nas linhas RESUMO

Extrai “Total no dia: N” do texto de observação

#### Tipo de execução

Execução manual (botão)

#### Observações técnicas

Controla ScreenUpdating, EnableEvents, Calculation para performance

Garante retorno ao estado anterior ao finalizar, mesmo em erro

Mantém marco de processamento (MARCO) para retomar de onde parou.

# Mapa rápido de dependências (VBA)

ThisWorkbook
→ modPQPowerQuery

modExecucao
→ modPQPowerQuery (opcional)
→ modAplicarFormulasPendencias
→ modLayoutPadrao (via rotina wrapper)
→ modSincronizarContatos

modSincronizarContatos
→ tblHistoricoBitrix (Histórico Bitrix) + tabelas Final/Manual/Correção
