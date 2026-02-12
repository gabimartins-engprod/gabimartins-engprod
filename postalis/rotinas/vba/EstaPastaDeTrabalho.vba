Option Explicit

' =====================================================================================
' MÓDULO: EstaPastaDeTrabalho (ThisWorkbook)
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   - Configurar comportamento do Power Query ao abrir o arquivo
' OBS:
'   - Execução manual via botão AtualizarTudo_E_Sincronizar (modExecucao)
' DEPENDÊNCIAS:
'   - PQ_BackgroundOff (modPQPowerQuery)
' VERSÃO: 12/02/2026
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    Application.EnableEvents = True
    PQ_BackgroundOff
    On Error GoTo 0
End Sub
