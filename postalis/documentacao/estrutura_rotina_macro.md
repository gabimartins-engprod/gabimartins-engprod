# Estrutura da Rotina — Macros (VBA)

Sistema: Rotina_Dados – Postalis  
Responsável: Gabriela Martins – Postalis – GBE/CCB  
Data de referência: 27/02/2026  

---

## 🎯 Objetivo

Documentar a macro principal responsável por orquestrar toda a execução da rotina:

- Atualização das consultas Power Query
- Sincronização de bases internas
- Geração de Log Histórico
- Correções estruturais pós-processamento
- Atualização de Tabelas Dinâmicas

---

## 🔘 Macro Principal

### AtualizarTudo_E_Sincronizar  
Localização: `modExecucao`

### Função
Orquestrar toda a execução da rotina do arquivo a partir de um único botão.

Fluxo controlado por:
- Trava anti-dupla execução (`isRunning`)
- Controle run-once por etapa (`Scripting.Dictionary`)
- Espera explícita do Power Query (`CalculateUntilAsyncQueriesDone`)

---

## 🔄 Fluxo de Execução

1. Desativa refresh em background (se módulo existir)
2. Executa `RefreshAll`
3. Aguarda finalização do Power Query
4. Executa sincronizações internas:
   - modAplicarFormulasPendencias
   - modSincronizarContatos
5. Executa backfill de datas pendentes
6. Gera Log Histórico baseado no Histórico_Bitrix
7. Recalcula e corrige TotalDia (apenas linhas RESUMO)
8. Atualiza Tabelas Dinâmicas

---

## 🧠 Dependências Críticas

### Tabelas obrigatórias

- `tblLogIDs` (aba Log_Sincronizacao)
- `tblHistoricoBitrix` (aba Histórico Bitrix)

### Colunas obrigatórias no Histórico

- ID (texto)
- DataArquivo (date)
- DataHoraExtracao (datetime opcional)

---

## ⚠️ Regras Técnicas Importantes

- DataArquivo é a referência oficial do dia.
- DataHoraExtracao é usada apenas para identificar a última extração válida do dia.
- IDs são tratados como texto.
- Apenas dias com extração são considerados no comparativo.
- O fix do TotalDia escreve valores somente nas linhas RESUMO.

---

## 🔐 Portabilidade

Para reutilização em outro ambiente:

1. Garantir que os nomes das tabelas sejam idênticos.
2. Manter as colunas obrigatórias.
3. Preservar constantes do topo do módulo.
4. Garantir que o Power Query não esteja rodando em Background.

---

## 📌 Versão

Versão atual da macro: 25/02/2026  
Patches aplicados:
- Robustez DataArquivo
- Fix Primeiro Dia com Extração
- Controle run-once
- Correção TotalDia via VBA
