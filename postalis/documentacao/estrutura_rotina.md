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
