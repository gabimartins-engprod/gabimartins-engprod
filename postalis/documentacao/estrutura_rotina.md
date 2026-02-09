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

### Histórico_Bitrix - VERSÃO: Atualizado em 09/02/2026

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

