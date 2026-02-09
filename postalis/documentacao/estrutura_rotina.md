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

## Bruta_Bitrix

Consulta base do sistema **Rotina_Dados – Postalis**.  
Responsável por ler a extração mais recente do Bitrix a partir de uma pasta local e padronizar os dados para uso nas demais consultas.

### Função
- Localizar automaticamente o arquivo de extração mais recente
- Identificar a aba ou tabela correta contendo os dados
- Normalizar nomes de colunas
- Garantir a existência de colunas obrigatórias
- Tipar campos principais
- Acrescentar metadados da extração

### Origem dos dados
- Pasta local contendo arquivos Excel de extração do Bitrix
- Prioridade para arquivos iniciados com `GT_PP_`
- Caso não existam, utiliza arquivos iniciados com `Extração Bitrix`

### Tipo de carregamento
- Apenas conexão (staging)
- Não carrega diretamente em planilha

### Colunas garantidas (contrato)
A consulta sempre entrega as seguintes colunas:

- ID  
- Parent task ID  
- Estágio  
- Tags  
- Status  
- Created on  
- Modified on  
- Task  
- Description  
- NomeArquivo  
- DataArquivo  

Caso alguma coluna não exista na origem, ela é criada com valores nulos.

### Observações técnicas
- O arquivo mais recente é definido pelo campo **Date modified**
- O caminho da pasta de extração é configurável (hardcoded nesta versão)
- Esta consulta serve como base para:
  - Correção_Automática
  - Registros_Contatos_Auto
  - Histórico_Bitrix

    ---

## Correção_Automática

Consulta que gera a base **limpa e padronizada** usada no restante do fluxo (ex.: Registros_Contatos_Auto).

### Função
- Recebe os dados da `Bruta_Bitrix`
- Garante a existência da coluna `Parent task ID`
- Mantém apenas colunas necessárias ao fluxo de contatos
- Aplica tipagem nas colunas de data (Created on, Modified on, DataArquivo)

### Origem dos dados
- Consulta `Bruta_Bitrix`

### Tipo de carregamento
- Carrega em planilha (tabela do Excel): **Correção_Automática**

### Colunas de saída (contrato)
- ID
- Parent task ID
- Estágio
- Tags
- Status
- Created on
- Modified on
- Task
- Description
- NomeArquivo
- DataArquivo

### Observações técnicas
- Serve como base para `Registros_Contatos_Auto` e demais consultas dependentes.
- Mudanças em nomes de colunas podem impactar o VBA (sincronização e fórmulas).

---

## Registros_Contatos_Auto

Consulta responsável por criar automaticamente a base de contatos a partir da Correção_Automática, consolidando informações contidas no texto do campo Task em uma estrutura única, padronizada e deduplicada.

‐
### Responsabilidades

Extrair Plano, Nome e Matrícula a partir do texto do campo Task;

Manter os campos de controle e rastreabilidade:

ID

Parent task ID

Created on;

Normalizar separadores textuais (-, –, —);

Remover duplicidades garantindo 1 linha por ID;

Manter sempre o registro mais recente com base no campo Created on.

‐
### Origem dos dados

Consulta Correção_Automática

‐
### Tipo de carregamento

Carrega em planilha do Excel

Tabela: Registros_Contatos_Auto

‐
### Colunas de saída (contrato)

ID

Parent task ID

Created on

Plano

Nome

Matrícula

‐
### Observações técnicas

A extração do Nome utiliza heurística baseada em tokens separados por “-”, exigindo presença de espaço, ausência de dígitos e exclusão por lista de palavras bloqueadas;

A extração da Matrícula identifica sequências numéricas entre 6 e 12 dígitos, retornando apenas os números;

A deduplicação é realizada por ID, preservando o registro mais recente conforme Created on.

### Histórico_Bitrix

Consulta responsável por consolidar o histórico diário das extrações do Bitrix, garantindo correspondência exata entre cada arquivo de extração e o respectivo dia.

**Principais responsabilidades:**
- Ler os arquivos de extração Bitrix a partir de pasta local configurada
- Considerar apenas arquivos Excel válidos (`.xlsx`, `.xls`, `.xlsm`)
- Priorizar arquivos iniciados com `GT_PP_`
- Selecionar automaticamente **um único arquivo por dia**, definido pelo campo *Data de modificação*
- Identificar dinamicamente a aba ou tabela que contém a coluna **ID**
- Normalizar e padronizar as principais colunas
- Adicionar os campos **NomeArquivo** e **DataArquivo**
- Consolidar histórico diário por ID, mantendo uma linha por ID por data

**Origem dos dados:**
- Pasta local contendo os arquivos Excel de extração do Bitrix
- Caminho configurado diretamente na consulta Power Query

**Tipo de carregamento:**
- Carrega em planilha (tabela `tblHistoricoBitrix`)
- Aba mantida oculta, utilizada para auditoria e validações pontuais

**Colunas entregues (contrato):**
- ID  
- Parent task ID  
- Estágio  
- Tags  
- Status  
- Created on  
- Modified on  
- Task  
- Description  
- NomeArquivo  
- DataArquivo  

> Caso alguma coluna não exista na origem, ela é criada com valores nulos.

**Observações técnicas:**
- O arquivo considerado por dia é definido pelo campo **Date modified**
- O histórico mantém consistência 1:1 entre extração diária e registros carregados
- Esta consulta serve como base para:
  - Registros_Contatos_Auto
  - Registros_Contatos_Final

---

