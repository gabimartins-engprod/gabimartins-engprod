Option Explicit

' =====================================================================================
' MÓDULO: ThisWorkbook (EstaPastaDeTrabalho)
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   - Ajustar configurações do Power Query na abertura do arquivo (desativar refresh em background),
'     garantindo previsibilidade na execução do botão Atualizar Base e Sincronizar.
'
' OBS:
'   - A execução da rotina principal é manual via botão AtualizarTudo_E_Sincronizar (modExecucao).
'
' DEPENDÊNCIAS:
'   - PQ_BackgroundOff (modPQPowerQuery)
'
' VERSÃO: 13/02/2026
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    PQ_BackgroundOff
End Sub
