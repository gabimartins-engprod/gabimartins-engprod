Option Explicit

' =====================================================================================
' MÓDULO: modPQPowerQuery
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   Garantir que as consultas do Power Query NÃO sejam atualizadas em background,
'   permitindo que o VBA controle o refresh de forma encadeada e previsível
'   (RefreshAll + CalculateUntilAsyncQueriesDone no modExecucao).
'
' ORIGEM DOS DADOS:
'   - ListObjects conectados a Power Query (QueryTable, quando aplicável)
'   - Conexões do Workbook (OLEDBConnection / ODBCConnection, quando existirem)
'
' DESTINO:
'   - Configurações de atualização (BackgroundQuery=False) aplicadas no próprio arquivo
'
' COMPORTAMENTO TÉCNICO:
'   - Percorre todas as tabelas (ListObjects) e, quando houver QueryTable:
'       * QueryTable.BackgroundQuery = False
'       * QueryTable.RefreshStyle = xlInsertDeleteCells
'   - Percorre todas as conexões do Workbook e tenta forçar BackgroundQuery=False
'     em OLEDBConnection e ODBCConnection (quando disponíveis)
'   - Erros são tratados com Debug.Print (não interrompe a rotina)
'
' MACROS PÚBLICAS:
'   - PQ_BackgroundOff: aplica BackgroundQuery=False em tabelas e conexões
'
' DEPENDÊNCIAS:
'   - Chamado por: ThisWorkbook.Workbook_Open (EstaPastaDeTrabalho)
'   - Usado por: modExecucao.AtualizarTudo_E_Sincronizar (etapa opcional)
'
' VERSÃO: 23/02/2026
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Public Sub PQ_BackgroundOff()
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If Not lo Is Nothing Then
                If HasQueryTable(lo) Then
                    lo.QueryTable.BackgroundQuery = False
                    lo.QueryTable.RefreshStyle = xlInsertDeleteCells
                End If
            End If
        Next lo
    Next ws

    ForceConnectionsBackgroundOff
    Exit Sub

TrataErro:
    Debug.Print "Erro em PQ_BackgroundOff: " & Err.Number & " - " & Err.Description
End Sub

Private Function HasQueryTable(ByVal lo As ListObject) As Boolean
    On Error GoTo Falha
    HasQueryTable = Not (lo.QueryTable Is Nothing)
    Exit Function
Falha:
    HasQueryTable = False
End Function

Private Sub ForceConnectionsBackgroundOff()
    On Error Resume Next

    Dim cn As WorkbookConnection
    For Each cn In ThisWorkbook.Connections
        If Not cn Is Nothing Then
            If Not cn.OLEDBConnection Is Nothing Then
                cn.OLEDBConnection.BackgroundQuery = False
            End If
            If Not cn.ODBCConnection Is Nothing Then
                cn.ODBCConnection.BackgroundQuery = False
            End If
        End If
    Next cn

    On Error GoTo 0
End Sub
