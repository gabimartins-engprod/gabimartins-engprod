# Macro: AtualizarTudo_E_Sincronizar

Módulo: modExecucao  
Sistema: Rotina_Dados – Postalis  

---

## 🎯 Finalidade

Executar toda a rotina de atualização, sincronização e geração de histórico do sistema em um único fluxo controlado.

---

## 🏗 Estrutura Interna

### Controle de Execução
- `isRunning`
- `gRanSteps` (Dictionary)
- RunStepOnce()

### Controle de Ambiente
- Desativa ScreenUpdating
- Desativa Events
- Controla Calculation manual/automática

---

## 🔄 Etapas Internas

### 1️⃣ Preparação
- modPQPowerQuery.PQ_BackgroundOff (opcional)

### 2️⃣ Atualização Power Query
- ThisWorkbook.RefreshAll
- Application.CalculateUntilAsyncQueriesDone

### 3️⃣ Sincronização Interna
- Rotina_Sincronizar_E_Recalcular
  → aplica layout padronizado e recalcula estruturas auxiliares
- Sincronizar_AteX_RegistrosContatos
  → sincroniza Registros_Contatos_Final com Correção_Automática
- Inclusão manual de novos IDs
  → Manual_IncluirNovosIDs_PeloFinal
  → Manual_IncluirIDs_PeloHistorico_0202
- Backfill pendente
  → Manual_Backfill_Rapido_0202

### 4️⃣ Geração do Log
- BuildLog_FromHistorico_FillMissingDays
- Comparativo entre dias
- Identificação de:
  - ENTROU
  - SAIU
  - RETORNOU

### 5️⃣ Correção Final
- FixTotalDia_IfExists
- Atualização de pivôs

---

## 📊 Estrutura do Log

Ações geradas:

- INICIO
- ENTROU
- RETORNOU
- SAIU
- RESUMO
- FIM
- MARCO

---

## 🧩 Integração com Power Query

Relaciona-se principalmente com:

- Histórico_Bitrix
- Correção_Automática
- Registros_Contatos_Auto
- Registros_Contatos_Final

A macro não altera consultas — apenas consome os resultados delas.

---

## 📌 Observação Técnica

O sistema foi projetado para:

- Evitar duplicidade por clique
- Manter baseline conceitual a partir de 01/02/2026
- Garantir consistência mesmo com dias sem extração
- Evitar cenário de “todos ENTROU” após lacunas

---

## 👩‍💻 Responsável

Gabriela Martins  
Postalis – GBE/CCB  
2026
