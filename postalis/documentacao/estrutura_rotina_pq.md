# Estrutura de Rotina — Exemplo

Documento criado para organizar a lógica geral de uma rotina administrativa,
sem utilização de dados reais.

## Objetivo
Descrever a sequência de atividades realizadas em uma rotina típica,
visando organização, padronização e futura melhoria de processo.

## Etapas gerais
1. Recebimento da demanda
2. Análise inicial
3. Tratamento da informação
4. Registro e controle
5. Finalização da atividade

## Observações
Documento de caráter acadêmico e organizacional.

---

### Bruta_Bitrix - VERSÃO: Atualizado em 09/02/2026

Consulta base do sistema **Rotina_Dados – Postalis**.

Responsável por localizar automaticamente o arquivo de extração mais recente do Bitrix a partir de uma pasta local e padronizar os dados para uso nas demais consultas.

#### Responsabilidades principais
- Localizar arquivos de extração Bitrix em pasta definida
- Priorizar arquivos iniciados com `GT_PP_`
- Selecionar automaticamente o arquivo mais recente (por data de modificação)
- Identificar a aba ou tabela correta contendo os dados
- Normalizar nomes de colunas principais (ID, Estágio, Status, datas, descrições)
- Garantir a existência de todas as colunas obrigatórias
- Adicionar metadados do arquivo (NomeArquivo, DataArquivo)

#### Origem dos dados
- Pasta local contendo arquivos Excel de extração do Bitrix
- Prioridade para arquivos iniciados com `GT_PP_`
- Alternativa para arquivos iniciados com `Extração Bitrix`

#### Tipo de carregamento
- Apenas conexão (staging)
- Não carrega diretamente em planilha

#### Observações técnicas
- O arquivo mais recente é definido pelo campo **Date modified**
- O caminho da pasta está definido no código nesta versão
- Recomenda-se parametrização do caminho para portabilidade
- Consulta base para:
  - Correção_Automática
  - Registros_Contatos_Auto
  - Histórico_Bitrix

    ---

### Correção_Automática - VERSÃO: Atualizado em 09/02/2026

Camada de tratamento automático da extração Bitrix.

- **Origem:** Bruta_Bitrix  
- **Finalidade:** Padronizar e limpar os dados brutos, gerando a base oficial utilizada pelas demais consultas do fluxo.  
- **Responsabilidades:**
  - Validar a existência da coluna essencial `ID`;
  - Garantir a coluna `Parent task ID`;
  - Selecionar apenas as colunas necessárias ao monitoramento de contatos;
  - Forçar tipagem consistente (IDs como texto e datas normalizadas);
  - Converter `DataArquivo` para tipo *date* de forma robusta.
- **Observação:**  
  Não possui qualquer lógica de histórico, comparação temporal ou cálculo de status. Atua exclusivamente como camada de staging.

Esta consulta deve ser utilizada como **fonte única** pelas consultas subsequentes, evitando consumo direto da Bruta_Bitrix.

---

### Registros_Contatos_Auto - VERSÃO: Atualizado em 09/02/2026

Camada de extração automática de informações textuais a partir da base tratada.

- **Origem:** Correção_Automática  
- **Finalidade:** Extrair automaticamente, a partir do campo `Task`, os campos `Plano`, `Nome` e `Matrícula`, mantendo `ID`, `Parent task ID` e `Created on`.  
- **Regras principais:**
  - Normalização do texto de `Task` para padronização de separadores;
  - Heurística para identificação de nomes próprios (tokens com espaço, sem dígitos e sem palavras bloqueadas);
  - Extração de matrícula por padrão numérico (6 a 12 dígitos);
  - Remoção de linhas com `ID` nulo ou vazio;
  - Deduplicação final garantindo 1 registro por `ID`, preservando o `Created on` mais recente.
- **Observação:**  
  Consulta intermediária, sem lógica de status, histórico ou controle temporal. Deve ser utilizada exclusivamente como base para `Registros_Contatos_Final`.
  
---

### Histórico_Bitrix - VERSÃO: Atualizado em 12/02/2026

Consulta responsável por consolidar o histórico diário das extrações do Bitrix,
garantindo rastreabilidade completa por ID e controle correto de movimentações
(ENTROU / SAIU / RETORNOU) no log de sincronização.

**Principais responsabilidades:**

• Ler os arquivos de remoção do Bitrix a partir da pasta local configurada
• Considerar apenas arquivos Excel válidos (.xlsx, .xls, .xlsm)
• Priorizar arquivos iniciados com GT_PP_
• Selecionar automaticamente um único arquivo por dia
• Identificar dinamicamente a aba ou tabela que contém a coluna ID
• Normalizar e padronizar as principais colunas
• Adicionar os campos NomeArquivo, DataArquivo e DataHoraExtracao
• Consolidar histórico diário por ID, mantendo uma linha por ID por dia

**Regra oficial para múltiplas extrações no mesmo dia:**

Quando houver mais de uma extração no mesmo dia,
a consulta mantém todas as linhas no histórico,
porém a coluna DataHoraExtracao (DateTime real)
permite que o VBA considere apenas a ÚLTIMA extração do dia
para cálculo de movimentações.

**Origem dos dados:**

• Pasta local contendo os arquivos Excel de remoção do Bitrix
• Caminho configurado diretamente na consulta Power Query

**Tipo de carregamento:**

• Carrega em planilha (tabela tblHistoricoBitrix)
• Aba mantida oculta, utilizada como base oficial do log histórico

**Colunas entregues (contrato):**

EU IA
ID
ID da tarefa principal
Estágio
Etiquetas
Status
Criado em
Modificado em
Tarefa
Descrição
NomeArquivo
DataArquivo
DataHoraExtracao  ← (DateTime oficial da extração)

Caso alguma coluna não exista na origem, ela é criada com valores nulos.

**Observações técnicas:**

• A coluna DataHoraExtracao é derivada do NomeArquivo e contém data + hora real da extração.
• Esta coluna é utilizada pelo VBA para:
   - Identificar a última extração diária
   - Calcular corretamente ENTROU / SAIU / RETORNOU
   - Reconstruir backlog de dias não executados
• A comparação diária não depende mais apenas da DataArquivo.

**Esta consulta serve como base para:**

• Registros_Contatos_Auto
• Registros_Contatos_Final
• Log_Sincronizacao (via VBA - modExecucao)

---

### Registros_Contatos_Manual - VERSÃO: Atualizado em 09/02/2026

Consulta destinada à leitura da base manual persistente de contatos.

**Finalidade**  
Expor a tabela de preenchimento manual com tipagem consistente, servindo como base auxiliar para:
- merge com a base automática (`Registros_Contatos_Final`);
- uso em rotinas VBA;
- cálculo posterior de métricas (ex.: último contato, dias sem contato).

**Características principais**
- Não altera dados.
- Não recria a tabela.
- Não remove nem insere registros.
- Mantém 1 linha por ID (controle humano).
- Blindada contra ausência ou renomeação de colunas.

**Regras de tipagem**
- `Created on`: datetime (data e hora).
- `Data_1` a `Data_5`: date (somente data).
- Campos textuais (ID, telefone, observações, responsáveis, e-mail, etc.): text.

**Fonte de dados**
- Tabela Excel `Registros_Contatos_Manual` (ThisWorkbook).

**Tipo de carga**
- Leitura direta local, sem dependências externas.

**Observação**
Esta consulta representa a camada de dados manuais persistentes do sistema e não deve sofrer transformações que impliquem perda de informação.

---

### Registros_Contatos_Final - VERSÃO: Atualizado em 09/02/2026

Consulta principal de consolidação da Rotina_Dados – Postalis.

**Finalidade**  
Unificar a base automática de contatos com os dados manuais persistentes, produzindo a visão final resumida utilizada para acompanhamento operacional e métricas.

**Fluxo lógico**
- Base automática (`Registros_Contatos_Auto`) como referência de IDs ativos.
- Base manual (`Registros_Contatos_Manual`) como fonte de histórico humano de contatos.
- Merge por `ID` com proteção para manter apenas uma linha manual por ID.
- Aplicação da regra oficial de “Último Contato”.

**Regra de Último Contato**
- Considera exclusivamente `Data_1` a `Data_5`.
- Último contato = maior data preenchida.
- Em caso de empate, prevalece o maior índice (5 > 4 > … > 1).
- OBS e Responsável são associados à mesma posição da data escolhida.

**Campos derivados**
- `Data` (último contato)
- `OBS` (observação do último contato)
- `Responsável` (responsável do último contato)
- `Qtd. tentativas` (quantidade de datas preenchidas entre Data_1 e Data_5)
- `Dias sem contato` (diferença em dias entre hoje e o último contato)

**Características**
- Não altera dados de origem.
- Não recria nem remove registros.
- Mantém uma linha por ID na saída.
- Preparada para integração com rotinas VBA (Dentro/Fora, Datas históricas, logs).

**Fonte de dados**
- Consultas internas do Power Query.

**Tipo de carga**
- Consolidação local (Power Query).

---

### Monitoramento_Extracao - VERSÃO: Atualizado em 09/02/2026

Consulta de apoio utilizada para validação técnica e análise histórica do ciclo dos IDs.

**Finalidade**  
Fornecer uma visão consolidada do universo histórico de IDs, permitindo identificar:
- IDs atualmente **Dentro** da extração;
- IDs **Fora** da extração (com data da última aparição);
- entradas recentes, com base na **primeira aparição registrada no histórico**.

**Fluxo lógico**
- Leitura do histórico consolidado (`Histórico_Bitrix`);
- Cálculo da primeira e da última aparição por ID;
- Comparação com a extração atual (`Registros_Contatos_Auto`);
- Classificação do status do ID (Dentro/Fora).

**Campos principais**
- `ID`
- `Primeira Aparição`
- `Última Aparição`
- `Status` (Dentro da extração / Fora da extração)
- `Data Saída` (aplicável apenas a IDs fora da extração)
- `NomeArquivo (Último)` (campo auxiliar para auditoria)

**Características**
- Consulta de caráter **técnico e analítico**.
- Não utilizada como rotina operacional.
- Não altera dados de origem.
- Não controla reentradas nem ciclos operacionais diários.

**Fonte de dados**
- Consultas internas do Power Query:
  - `Histórico_Bitrix`
  - `Registros_Contatos_Auto`

**Tipo de carga**
- Consolidação local (Power Query).

