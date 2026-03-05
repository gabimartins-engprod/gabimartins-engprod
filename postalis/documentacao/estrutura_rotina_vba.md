# Estrutura de Rotina — VBA (Rotina_Dados – Postalis)

Documento criado para organizar a lógica geral do **VBA** da rotina administrativa **Rotina_Dados – Postalis**, descrevendo responsabilidades, dependências e sequência de execução, **sem uso de dados reais.**

## Objetivo

Documentar a arquitetura do **fluxo VBA**, deixando claro:

- quais módulos existem;

- qual a responsabilidade de cada um;

- como eles se conectam ao Power Query e às tabelas;

- como o botão único executa a rotina ponta a ponta.

## Etapas gerais (fluxo VBA)

**1.** Abertura do arquivo (Workbook_Open)

**2.** Preparação do Power Query (BackgroundQuery = False)

**3.** Atualização das consultas (RefreshAll + Wait)

**4.** Aplicação de layout padronizado (sem fórmulas na arquitetura atual)

**5.** Sincronização da base final (inclusão/atualização de IDs + Data Entrada)

**6.** Sincronização do Manual (inclusão de IDs + backfill pendentes)

**7.** Geração/atualização do Log_Sincronizacao (dias com e sem extração)

**8.** Recalcular + correções finais (TotalDia)

**9.** Atualização de pivôs

**10.** Finalização controlada

## Observações

- Documento de caráter organizacional e técnico.

- A rotina principal é **manual** via botão **AtualizarTudo_E_Sincronizar** (módulo modExecucao).

- A lógica diária de comparação é baseada em **DataArquivo (date)** como “dia oficial” e **DataHoraExtracao (datetime)** apenas para selecionar a **última extração do dia**, quando disponível.

 ---

### ThisWorkbook (EstaPastaDeTrabalho) — VERSÃO: 13/02/2026

Módulo de evento do Excel, responsável por **preparar o ambiente** quando o arquivo é aberto.

#### Finalidade

Garantir previsibilidade na execução do fluxo VBA, chamando a rotina que desativa execução em background das consultas do Power Query.

#### Responsabilidades principais

- Executar automaticamente ao abrir o arquivo (Workbook_Open)

- Chamar PQ_BackgroundOff para desligar refresh em background

#### Dependências

- modPQPowerQuery.PQ_BackgroundOff

#### Tipo de execução

- Automática (evento do Excel)

#### Observações técnicas

- Usa On Error Resume Next para **não impedir a abertura** do arquivo caso haja alguma divergência/erro.

 ---

### modPQPowerQuery — VERSÃO: 23/02/2026

Módulo responsável por **forçar o Power Query a não rodar em background**, permitindo que o VBA controle o encadeamento de atualizações.

#### Finalidade

Evitar execução assíncrona do Power Query, garantindo que o VBA consiga fazer:

- RefreshAll

- CalculateUntilAsyncQueriesDone
com consistência.

#### Responsabilidades principais

- Percorrer todas as abas e ListObjects

- Se a tabela tiver QueryTable, setar:

     → BackgroundQuery = False

     → RefreshStyle = xlInsertDeleteCells

- Percorrer ThisWorkbook.Connections e, quando aplicável:

     → OLEDBConnection.BackgroundQuery = False

     → ODBCConnection.BackgroundQuery = False

#### Origem/entrada

- Estruturas do Excel (ListObjects, QueryTables e Connections)

#### Tipo de execução

- Preparação técnica (configuração do ambiente)

#### Observações técnicas

- Tratamento de erro com Debug.Print para diagnóstico sem travar a rotina.

 ---

### modLayoutPadrao — VERSÃO: 23/02/2026

Módulo de padronização visual e performance do layout, com AutoFit leve.

#### Finalidade

Aplicar layout padronizado nas abas principais sem o custo do AutoFit completo na planilha inteira.

#### Abas alvo

- Extração Bitrix

- Registros Contatos

- Registros Contatos Manual

#### Responsabilidades principais

- Tipo de responsabilidade:
  
   → Camada exclusivamente visual (não interfere na lógica de dados)

- Ajustar altura de linhas:

     → Linha 1: altura maior (cabeçalho)

     → Linha 2: altura intermediária

     → Linhas 3 até o final: altura padrão

- Aplicar AutoFit **apenas nas colunas das tabelas**, usando **amostra** de linhas

#### Controle de performance

- AutoFit roda **uma vez por dia**, salvo execução forçada

- Usa o Name do workbook:

     → RD_LastAutoFitDate

#### Tipo de execução

- Ajuste visual pós-processamento (layout)

#### Observações técnicas

- AutoFitListObject_Sample: pega Header + N primeiras linhas (amostragem) e faz AutoFit apenas nessas colunas.

 ---

### modAplicarFormulasPendencias — VERSÃO: 03/03/2026

Módulo mantido por compatibilidade com o fluxo principal do sistema (modExecucao), atuando como etapa intermediária para padronização visual e encadeamento da rotina.

**Finalidade**
Na arquitetura atual, este módulo NÃO aplica fórmulas na tabela Correção_Automática e NÃO cria colunas.
Ele permanece ativo para:
- manter compatibilidade com chamadas existentes do botão principal;
- aplicar layout padronizado ao final do refresh (via modLayoutPadrao).

**Responsabilidades principais**
- Garantir que nenhuma coluna estrutural seja criada automaticamente na Correção_Automática.
- Servir como “ponte” de compatibilidade para o modExecucao.
- Aplicar layout padronizado após atualização.

**Integração com layout**
Wrapper Rotina_Sincronizar_E_Recalcular chama:
→ AplicarFormulasPendencias (vazio por design na arquitetura atual)
→ modLayoutPadrao.AplicarLayoutPadraoTabelas

**Tipo de execução**
Etapa de pós-processamento (layout) + compatibilidade do fluxo VBA.

**Observações técnicas**
- O módulo está “blindado” para não alterar a Correção_Automática.
- Se no futuro for necessário espelhar campos, isso deve ser feito SOMENTE em colunas já existentes (planejado previamente).

---

### modSincronizarContatos — VERSÃO: 03/03/2026

Módulo central de sincronização e manutenção de bases de contato, responsável por manter consistência entre a base automática (Correção_Automática), a base final (Registros_Contatos_Final) e a base manual (Registros_Contatos_Manual), aplicando regra oficial de ciclo baseada em DataArquivo (Histórico Bitrix).

**Finalidade**
1) Sincronizar Registros_Contatos_Final com Correção_Automática (inclusão e atualização de IDs).
2) Manter e alimentar Registros_Contatos_Manual com IDs e dados operacionais (Status, Data Saída, Data Retorno e backfill).

**Tabelas**
- Correção_Automática (base automática tratada)
- Registros_Contatos_Final (base oficial consolidada — mantém apenas Data Entrada)
- Registros_Contatos_Manual (base operacional — controla Status, Data Saída e Data Retorno)
- Histórico: tblHistoricoBitrix (aba Histórico Bitrix)

**Baseline e regras**
- BASELINE OFICIAL: 01/02/2026
- Backfill/ciclo percorrem somente dias com extração (dias existentes em DataArquivo no Histórico).
- DataHoraExtracao é opcional e, quando existe, serve para selecionar “última foto do dia”.

**Responsabilidades principais (por rotina)**

**(A)** Sincronizar_AteX_RegistrosContatos
Sincroniza Final com Correção_Automática garantindo:
- inclusão de novos IDs vindos da extração;
- atualização de Nome e Matrícula quando disponíveis;
- preenchimento de Data Entrada somente quando vazia (dataRef = Dia do Painel);
- NÃO criação de colunas estruturais (evita duplicidades).

⚠️ **Importante:**
A tabela Final mantém apenas Data Entrada.
Status, Data Saída e Data Retorno são controlados exclusivamente na tabela Manual.

**(B)** Manual_IncluirNovosIDs_PeloFinal
Inclui IDs que existem no Final, mas não existem no Manual.
Copia campos-chave: Parent task ID, Created on, Matrícula e Nome.

**(C)** Manual_IncluirIDs_PeloHistorico_0202
Inclui IDs novos diretamente do Histórico desde o baseline.
Critério: pega a linha mais recente por ID (maior DataArquivo; desempate por DataHoraExtracao quando existir).

**(D)** Manual_Backfill_Pendentes_0202 (Backfill leve – somente pendentes)
Verifica pendências de dados no Manual:
→ Data Entrada
→ Data Saída
→ Data Retorno
Regras:
- Limpa datas inválidas (dias que não são dia de extração).
- Calcula saída/retorno percorrendo apenas extrDates (dias com extração).

**Tipo de execução**
Execução por etapas, chamada pelo modExecucao.

**Observações técnicas**
- IDs tratados como texto (NumberFormat="@").
- Usa Scripting.Dictionary para performance (sets e presença diária).
- Preenchimento limitado a pendentes (mais leve e seguro).
- O sistema nunca remove registros inseridos manualmente (base persistente).

 ---

### modExecucao — VERSÃO: 25/02/2026

Módulo orquestrador do sistema (botão único), responsável por rodar a rotina completa com proteção anti-duplicidade.

#### Finalidade

Executar o fluxo ponta a ponta: atualizar consultas, aplicar regras, sincronizar tabelas e gerar log histórico.

#### Botão único

- AtualizarTudo_E_Sincronizar

#### Etapas executadas (sequência)

**1.** PQ_BackgroundOff (opcional)

**2.** RefreshAllAndWait (Power Query)

**3.** modAplicarFormulasPendencias.Rotina_Sincronizar_E_Recalcular

**4.** modSincronizarContatos.Sincronizar_AteX_RegistrosContatos

**5.** Manual_IncluirNovosIDs_PeloFinal + Manual_IncluirIDs_PeloHistorico_0202

**6.** Manual_Backfill_Rapido_0202

**7.** BuildLog_FromHistorico_FillMissingDays (desde 01/02/2026)

**8.** Recalcular + FixTotalDia_IfExists

**9.** Atualizar tabelas dinâmicas (RefreshAllPivots)

#### Regras críticas do Log (Histórico)

- Referência de “dia”: **DataArquivo**

- Seleção da “foto válida do dia”:

     → se existe DataHoraExtracao, usa apenas a maior daquele dia

     → se não existe, usa todas as linhas do dia por DataArquivo

- Gera blocos por dia:

     → INICIO

     → (RETORNOU / SAIU / ENTROU) quando aplicável

     → RESUMO

     → FIM

- Para dias sem extração:

     → registra RESUMO com alerta para verificação

#### Anti-duplicidade por clique (Run once por etapa)

- Cache em dictionary gRanSteps

- Cada etapa roda no máximo uma vez por execução do botão

#### Correção TotalDia

- FixTotalDia_IfExists preenche TotalDia **somente nas linhas RESUMO**

- Extrai “Total no dia: N” do texto de observação

#### Tipo de execução

- Execução manual (botão)

#### Observações técnicas

- Controla ScreenUpdating, EnableEvents, Calculation para performance

- Garante retorno ao estado anterior ao finalizar, mesmo em erro

- Mantém marco de processamento (MARCO) para retomar de onde parou.

### Controle de Ambiente

Durante a execução:

- Application.ScreenUpdating = False

- Application.EnableEvents = False

- Application.Calculation = xlCalculationManual

E restaurado ao final, inclusive em erro.

---

# Mapa rápido de dependências (VBA)

- **ThisWorkbook**
     → modPQPowerQuery

- **modExecucao**
     → modPQPowerQuery (opcional)
     → modAplicarFormulasPendencias
     → modLayoutPadrao (via rotina wrapper)
     → modSincronizarContatos

- **modSincronizarContatos**
     → tblHistoricoBitrix (Histórico Bitrix) + tabelas Final/Manual/Correção
