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

### Histórico_Bitrix - VERSÃO: Atualizado em 13/02/2026

Consulta responsável por consolidar o histórico diário das extrações do Bitrix, garantindo correspondência exata entre cada arquivo de extração e o respectivo dia.

**🔎 Principais responsabilidades:**

- Ler arquivos Excel válidos (.xlsx, .xls, .xlsm) na pasta configurada
- Ignorar arquivos temporários (~$)
- Priorizar arquivos iniciados com GT_PP_
- Selecionar automaticamente um único arquivo por dia, definido por:
   DataArquivo (Date modified convertido para date)
   Ordenação por Date modified (mais recente do dia)
- Identificar dinamicamente a aba ou tabela que contenha a coluna ID
- Padronizar nomes e tipos de colunas obrigatórias
- Garantir integridade das colunas contratuais
- Consolidar histórico mantendo 1 linha por ID por DataArquivo

**⏱ Controle de Extração (Novo – 13/02/2026):**

Para garantir a regra “última extração do dia”, a consulta agora:
- Extrai a informação de horário diretamente do NomeArquivo (ex: 06h16min)
- Cria as colunas auxiliares:
   ParteHora
   HoraNum
  MinutoNum
- Cria a coluna oficial:
   DataHoraExtracao (datetime real)
- Essa coluna é utilizada pelo VBA (modExecucao) para:
   Identificar a última extração do dia
   Construir o log histórico corretamente
   Permitir futura auditoria de múltiplas extrações no mesmo dia
Se o padrão do nome do arquivo for alterado, essa lógica deve ser revisada.

**📥 Origem dos dados:**

- Pasta local configurada diretamente na consulta (Folder.Contents)
- Não varre subpastas

**📤 Tipo de carregamento:**

- Carrega em planilha
- Tabela: tblHistoricoBitrix
- Aba: Histórico Bitrix
- Utilizada como base para:
   Registros_Contatos_Auto
   Registros_Contatos_Final
   Construção do Log_Sincronizacao

**📦 Colunas entregues (contrato):**

ID
Parent task ID
Estágio
Tags
Status
Created on
Modified on
Task
Description
NomeArquivo
DataArquivo
ParteHora
HoraNum
MinutoNum
DataHoraExtracao

Caso alguma coluna obrigatória não exista na origem, é criada com valores nulos.

**🛡 Observações técnicas:**

- Mantém histórico desde 02/02/2026
- Para cada DataArquivo mantém apenas o arquivo mais recente do dia
- A regra “última extração do dia” é baseada na DataHoraExtracao (datetime real)
- Compatível com múltiplas extrações no mesmo dia

---

### Registros_Contatos_Manual - VERSÃO: Atualizado em 12/05/2026

Consulta destinada à leitura da base manual persistente de contatos.

**Finalidade**

Expor a tabela de preenchimento manual com tipagem consistente, utilizada como base auxiliar para:
- mesclar com a base automática (Registros_Contatos_Final);
- uso em rotinas VBA;
- cálculo posterior de métricas (ex.: último contato, dias sem contato).

**Características principais**

- Não altera dados.
- Não recria a tabela.
- Não remove nem insere registros.
- Mantém 1 linha por ID (controle humano).
- Blindada contra ausência ou renomeação de colunas.
- Aplica tipagem apenas nas colunas existentes da tabela.

**Regras de tipagem**

**Campos principais**

- Created on → data e hora.
- ID e Parent task ID → texto.
- Matrícula e Nome → texto.

**Campos de controle do ciclo**

- Data Entrada
- Data Saída
- Data Retorno
  
  → tipo data.

- Status → texto.
- Tentativas_Realizadas → número inteiro (Int64.Type).

**Campos auxiliares**
- Observação Extra
- Telefone
- Contato_Extra

→ texto.

**Histórico de contatos:**

- Data_1 até Data_10 → tipo data.
- Tipo_1 até Tipo_10 → texto.
- OBS_1 até OBS_10 → texto.
- Responsável_1 até Responsável_10 → texto.

**Fonte de dados**

Tabela Excel Registros_Contatos_Manual (Excel.CurrentWorkbook).

**Tipo de carga**

- Leitura direta local, sem dependências externas.

**Observação**

Esta consulta representa a camada de dados manuais persistentes do sistema e não deve sofrer transformações que impliquem perda de informação.

A tipagem é aplicada dinamicamente apenas nas colunas existentes, garantindo robustez caso colunas sejam adicionadas, removidas ou renomeadas na planilha manual.

---

### Registros_Contatos_Final - VERSÃO: Atualizado em 03/03/2026

Consulta principal de consolidação da Rotina_Dados – Postalis.

**Finalidade**

Unificar a base automática de contatos com os dados manuais persistentes, produzindo a visão final resumida utilizada para acompanhamento operacional e análises.

**Fluxo lógico**

- Base automática (Registros_Contatos_Auto) como referência de IDs ativos.

- Base manual (Registros_Contatos_Manual) como fonte de histórico humano de contatos.

- Merge por ID com proteção para manter apenas uma linha manual por ID.

- Aplicação da regra oficial de identificação do Último Contato.

- Utilização do Dia do Painel para cálculo de métricas operacionais.

**Regra de Último Contato**

Considera exclusivamente as colunas:
- Data_1 até Data_10.
- Último contato = maior data preenchida.
- Em caso de empate, vence o maior índice (10 > 9 > … > 1).
- OBS e Responsável são associados à mesma posição da data escolhida.

**Campos derivados**

- Data (último contato)
- OBS (observação do último contato)
- Responsável (responsável pelo último contato)
- Qtd. tentativas
- (quantidade de datas preenchidas entre Data_1 e Data_10)
- Dias sem contato
- (diferença em dias entre o Dia do Painel e o último contato)

**Dia do Painel**

O Dia do Painel é definido como:
- max(DataArquivo)
- da consulta Correção_Automática.
- Essa data representa o dia efetivo da última extração Bitrix.
- Caso não exista DataArquivo disponível, o sistema utiliza o baseline oficial do ciclo operacional:
01/02/2026
- Esse baseline garante que a rotina nunca utilize a data do computador como fallback, preservando a consistência dos cálculos.

**Proteções aplicadas**

- Remoção preventiva da coluna Data Entrada da base automática antes do merge, evitando duplicidade de campos.
- Uso exclusivo da Data Entrada da base manual como referência real de entrada do ID no ciclo de acompanhamento.
- Proteção contra valores negativos no cálculo de Dias sem contato.

**Características**

- Não altera dados de origem.
- Não recria nem remove registros.
- Mantém uma linha por ID na saída.
- Preparada para integração com rotinas VBA (controle de ciclo, logs e sincronizações).

**Fonte de dados**

Consultas internas do Power Query:
- Correção_Automática
- Registros_Contatos_Auto
- Registros_Contatos_Manual

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

