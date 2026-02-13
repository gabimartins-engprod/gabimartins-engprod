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


