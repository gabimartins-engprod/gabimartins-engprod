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

 
