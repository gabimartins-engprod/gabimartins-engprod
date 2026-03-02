# Arquitetura de Fluxo — Rotina_Dados (Postalis)

Sistema: Rotina_Dados – Postalis  
Responsável: Gabriela Martins  
Data de referência: 2026  

---

# 🎯 Objetivo

Documentar o fluxo completo de processamento do sistema, desde a leitura das extrações Bitrix até a geração do Log Histórico e sincronização das bases de contato.

Este documento descreve a arquitetura lógica do pipeline de dados.

---

# 🧠 Visão Geral do Fluxo

Extração Bitrix (.xlsx)
        ↓
Power Query (ETL)
        ↓
Tabelas Estruturadas
        ↓
AtualizarTudo_E_Sincronizar
        ↓
Sincronização + Regras de Ciclo
        ↓
Log Histórico
        ↓
Backfill
        ↓
Pivôs / Visualização

---

# 🔹 Etapa 1 — Extração de Dados

## Origem
Arquivos Excel exportados do Bitrix.

## Regras aplicadas:
- Seleciona arquivo mais recente
- Normaliza nomes de colunas
- Garante colunas obrigatórias
- Padroniza tipos (ID como texto)
- Adiciona:
  - NomeArquivo
  - DataArquivo (date)
  - DataHoraExtracao (datetime quando aplicável)

---

# 🔹 Etapa 2 — Camada Power Query (ETL)

Local: `rotina/power_query`

Consultas principais:

- bruta_bitrix.pq
- correcao_automatica.pq
- registros_contatos_auto.pq
- registros_contatos_manual.pq
- registros_contatos_final.pq
- historico_bitrix.pq
- monitoramento_extracao.pq

## Funções da camada:
- Limpeza
- Padronização
- Extração de atributos
- Consolidação de histórico
- Preparação das tabelas consumidas pelo VBA

---

# 🔹 Etapa 3 — Orquestração VBA

Ponto de entrada:

AtualizarTudo_E_Sincronizar

Local: `rotina/vba/modExecucao`

### Responsabilidades:

1. Desativar BackgroundQuery (segurança de execução)
2. Executar RefreshAll
3. Aguardar término do Power Query
4. Sincronizar tabelas internas
5. Executar backfill (se necessário)
6. Gerar Log Histórico incremental
7. Corrigir TotalDia
8. Atualizar Tabelas Dinâmicas

---

# 🔹 Regras Estruturais do Sistema

## 📌 Regra 1 — Data Oficial do Dia

`DataArquivo` é a referência oficial do dia.

`DataHoraExtracao` é usada apenas para selecionar a última extração válida daquele dia.

---

## 📌 Regra 2 — Comparativo por Dias com Extração

O sistema:

- Percorre apenas dias que possuem extração
- Ignora lacunas como movimento válido
- Evita cenário falso de "todos ENTROU"

---

## 📌 Regra 3 — Baseline Conceitual

Data de início do ciclo:

01/02/2026

Reconstruções históricas respeitam essa data como ponto inicial de ciclo.

---

## 📌 Regra 4 — Controle de Execução

- Controle anti-dupla execução (isRunning)
- Execução "run once" por etapa
- CalculateUntilAsyncQueriesDone para previsibilidade

---

# 🔹 Etapa 4 — Log Histórico

Tabela: `tblLogIDs`

Ações geradas:

- INICIO
- ENTROU
- RETORNOU
- SAIU
- RESUMO
- FIM
- MARCO

### Particularidade:

`TotalDia` é recalculado via VBA apenas nas linhas RESUMO.

---

# 🔹 Etapa 5 — Backfill Inteligente

Executado por:

- Manual_Backfill_Rapido_0202
- Manual_Backfill_Pendentes_0202

Regras:
- Considera somente dias com extração
- Preenche datas de entrada/saída/retorno
- Mantém integridade histórica

---

# 🧩 Integração Entre Camadas

| Camada | Responsabilidade |
|--------|-----------------|
| Power Query | ETL e estruturação |
| VBA | Orquestração e regras de negócio |
| Log | Auditoria histórica |
| Pivôs | Visualização |

---

# 📈 Arquitetura Conceitual

Este sistema pode ser interpretado como:

- Um pipeline ETL local
- Com camada de orquestração
- Com reconciliação histórica
- Com controle incremental de estado

---

# 🔐 Segurança e Integridade

- IDs tratados como texto
- Sem dependência de banco externo
- Histórico incremental controlado
- Evita duplicidade por clique
- Tratamento explícito de lacunas

---

# 📌 Conclusão

O sistema Rotina_Dados funciona como:

> Um pipeline ETL incremental local, com orquestração em VBA,
controle de estado por ciclo e reconciliação histórica baseada em extrações reais.

Arquitetura modular e separada por responsabilidade.

---

## 🧭 Diagrama do Fluxo (Mermaid)

<img width="2567" height="3224" alt="mermaid-diagram" src="https://github.com/user-attachments/assets/7e27fd08-6e3e-422c-a57f-180276e72593" />

Obs.: Código - Diagrama do Fluxo
flowchart TD
  A[Extração Bitrix<br/>(Arquivos .xlsx)] --> B[Power Query (ETL)<br/>Limpeza • Padronização • Consolidação]
  B --> C[Tabelas Estruturadas no Excel<br/>Histórico_Bitrix • Correção_Automática • Registros_Contatos_*]
  C --> D[Macro Principal<br/>AtualizarTudo_E_Sincronizar<br/>(modExecucao)]
  
  D --> E[Controle de Execução<br/>isRunning • RunStepOnce • Ambiente Excel]
  E --> F[Atualização PQ<br/>RefreshAll + CalculateUntilAsyncQueriesDone]
  F --> G[Sincronização<br/>modSincronizarContatos]
  G --> H[Aplicar Fórmulas<br/>modAplicarFormulasPendencias]
  H --> I[Geração Log Histórico<br/>tblLogIDs<br/>ENTROU/SAIU/RETORNOU + RESUMO]
  I --> J[Backfill (quando necessário)<br/>Manual_Backfill_*]
  J --> K[Correções finais<br/>FixTotalDia (somente RESUMO)]
  K --> L[Atualização de Pivôs<br/>Visualizações / Indicadores]

  %% Regras estruturais
  R1{{Regra: DataArquivo = dia oficial}} --- I
  R2{{Regra: DataHoraExtracao<br/>apenas última extração do dia}} --- F
  R3{{Regra: comparar apenas dias com extração}} --- I
  R4{{Baseline conceitual<br/>01/02/2026}} --- I
