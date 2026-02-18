## ThisWorkbook (EstaPastaDeTrabalho) - Atualizado em 13/02/2026

**Finalidade:** Ajustar configurações do Power Query ao abrir o arquivo, desativando a atualização em segundo plano (background) para garantir previsibilidade na execução da rotina principal via botão.

### Resumo rápido
Ao abrir o arquivo `Rotina_Dados`, executa `PQ_BackgroundOff` para desabilitar refresh em background do Power Query, evitando comportamentos assíncronos durante a atualização.

### Tipo de carga
- **Automática:** evento `Workbook_Open`

### Entradas (lê/usa)
- Configurações do Excel (Application)
- Procedure `PQ_BackgroundOff` (módulo `modPQPowerQuery`)

### Saídas (escreve/impacta)
- Configuração de Power Query: desabilita refresh em segundo plano (background)
- Garante `Application.EnableEvents = True` na abertura

### Dependências
- `modPQPowerQuery.PQ_BackgroundOff`

### Observações de portabilidade
- Se o nome do procedimento `PQ_BackgroundOff` ou do módulo `modPQPowerQuery` mudar, atualizar a chamada no `Workbook_Open`.
- Mantém tratamento de erro simples para não travar a abertura do arquivo caso o procedimento não esteja disponível.

 ---

 ## modExecucao - Atualizado em 13/02/2026

**Finalidade:** Orquestrar a execução do botão único **AtualizarTudo_E_Sincronizar**, garantindo atualização do Power Query, sincronização interna e geração do **Log Histórico diário** (por dia) a partir do Histórico Bitrix.

### Resumo rápido
Executa a sequência: (1) RefreshAll do Power Query, (2) rotina interna de sincronização da base, (3) refresh das tabelas dinâmicas, e (4) geração/atualização do log histórico diário (tblLogIDs), preenchendo dias faltantes com base na última extração de cada dia (DataHoraExtracao).

### Tipo de carga
- **Manual:** botão `AtualizarTudo_E_Sincronizar`

### Entradas (lê/usa)
- `Histórico Bitrix` → tabela `tblHistoricoBitrix`
  - Colunas obrigatórias: `ID`, `DataHoraExtracao` (DateTime)
- `Log_Sincronizacao` → tabela `tblLogIDs` (para identificar dias já processados via RESUMO/MARCO)

### Saídas (escreve/impacta)
- `Log_Sincronizacao` → tabela `tblLogIDs`
  - Registra blocos por dia: `INICIO`, (`ENTROU/SAIU/RETORNOU` se houver), `RESUMO`, `FIM`
  - Registra `MARCO` com a última data processada (controle incremental)
- Atualiza tabelas dinâmicas do arquivo (`PivotTables.RefreshTable`)

### Regras críticas
- **START_DATE:** considera dados a partir de `02/02/2026`.
- **Primeiro dia (START_DATE):** não realiza comparativo (sem dia anterior); grava apenas `INICIO + RESUMO (Total) + FIM`.
- **Mais de uma extração no dia:** utiliza a **última extração** (maior `DataHoraExtracao`) para montar o conjunto diário.
- **Dia sem extração:** grava bloco de aviso (`INICIO/RESUMO/FIM`) sem movimentações.
- **Anti-duplicação:** trava por dia baseado em `RESUMO` (DataArquivo) já existente no log.

### Dependências
- `Rotina_Sincronizar_E_Recalcular` (módulo externo da rotina)
- Tabelas devem existir com os nomes:
  - `Log_Sincronizacao` / `tblLogIDs`
  - `Histórico Bitrix` / `tblHistoricoBitrix`

### Observações de portabilidade
- Alterações de nomes de abas/tabelas/colunas exigem ajuste das constantes (`SH_LOG`, `TBL_LOG`, `SH_HIST`, `TB_HIST`, `COL_DATAHORA_EXTRACAO`).
- A coluna `DataHoraExtracao` deve ser DateTime válido para cálculo correto da “última extração do dia”.

---

## modSincronizarContatos - Atualizado em 18/02/2026

**Finalidade:**  
Sincronizar a base `Registros_Contatos_Final` com a consulta `Correção_Automática`, controlando o ciclo dos IDs (Status, Data Entrada e Data Saída) e mantendo a base `Registros_Contatos_Manual` atualizada sem perder informações inseridas manualmente.

### Resumo rápido
Compara diariamente os IDs da extração atual (`Correção_Automática`) com a tabela `Registros_Contatos_Final`, executando:

- Inclusão automática de novos IDs
- Atualização do Status:
  - "Dentro da extração"
  - "Fora da extração"
- Controle do ciclo:
  - Data Entrada (primeira presença ou retorno)
  - Data Saída (quando deixa de aparecer)
- Sincronização da aba manual sem apagar histórico

As datas são gravadas como **Date (sem hora)** para padronização e estabilidade em filtros, BI e análises.

### Tipo de carga
- Manual (executado via botão `AtualizarTudo_E_Sincronizar`)
- Encadeado dentro da rotina principal (modExecucao)

### Entradas (lê/usa)
- Tabela: `Correção_Automática`
- Tabela: `Registros_Contatos_Final`
- Tabela: `Registros_Contatos_Manual`

### Saídas (escreve/impacta)

🔹 `Registros_Contatos_Final`
- Status
- Data Entrada (Date, sem hora)
- Data Saída (Date, sem hora)
- Inclusão de novos IDs

🔹 `Registros_Contatos_Manual`
- Atualiza colunas automáticas (ID, Nome, Matrícula, Parent task ID, Created on)
- Preserva 100% das colunas manuais

### Dependências
- `GetTableByName`
- `EnsureListColumn`
- `ReadCell`
- `WriteCell`
- `IsEmptyOrNull`
- Chamado por `modExecucao`

### Observações de portabilidade
- As colunas "Status", "Data Entrada" e "Data Saída" são garantidas via VBA, mesmo que sejam removidas manualmente.
- Datas são gravadas usando `Date` (sem hora), evitando conflitos com filtros, agrupamentos e Power BI.
- Este módulo não grava log histórico — o log oficial é gerado exclusivamente no `modExecucao`.

---

## modPQPowerQuery - Atualizado em 13/02/2026

**Finalidade:** Garantir que o refresh do Power Query não rode em segundo plano (background), permitindo execução síncrona e previsível da rotina principal (RefreshAll + CalculateUntilAsyncQueriesDone) antes das etapas de VBA.

### Resumo rápido
Executa `PQ_BackgroundOff` para desativar refresh em background:
1) Nas tabelas carregadas em planilha (`ListObject` com `QueryTable`), define `BackgroundQuery = False` e `RefreshStyle = xlInsertDeleteCells`.
2) Nas conexões do arquivo (quando suportam `OLEDBConnection`), tenta definir `BackgroundQuery = False`.

Inclui ainda:
- `HookPQTables` (compatibilidade / reservado; atualmente vazio)
- `ReHookPQ` (executa o hook e exibe mensagem informativa)

### Tipo de carga
- **Automática:** chamado por `ThisWorkbook.Workbook_Open`
- **Manual:** pode ser executado diretamente (se necessário)

### Entradas (lê/usa)
- Todas as planilhas do `ThisWorkbook`
- `ListObjects` e suas propriedades (`QueryTable`, quando existir)
- `ThisWorkbook.Connections` (conexões do arquivo) e `OLEDBConnection` (quando existir)

### Saídas (escreve/impacta)
- Define `QueryTable.BackgroundQuery = False` (quando aplicável)
- Define `QueryTable.RefreshStyle = xlInsertDeleteCells` (quando aplicável)
- Define `OLEDBConnection.BackgroundQuery = False` (quando aplicável)
- Reduz risco de comportamento assíncrono durante a execução do botão principal

### Dependências
- Chamado por `ThisWorkbook.Workbook_Open`

### Observações de portabilidade
- Atua somente onde há `QueryTable` (tabelas carregadas na planilha). Se a consulta não estiver carregada como tabela, pode não ser afetada.
- A etapa de conexões depende do tipo de conexão: só aplica quando existir `OLEDBConnection`.
- Em caso de erro, o procedimento exibe mensagem (`MsgBox`) mas não deve impedir a abertura do arquivo.
