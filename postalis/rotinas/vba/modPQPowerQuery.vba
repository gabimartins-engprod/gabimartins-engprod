Option Explicit

' =====================================================================================
' MÓDULO: modPQPowerQuery
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   Ajustar o comportamento de refresh do Power Query para NÃO rodar em background,
'   permitindo sincronização encadeada com VBA (RefreshAll + CalculateUntilAsyncQueriesDone).
'
' COMO FUNCIONA:
'   (1) Força BackgroundQuery = False nas QueryTables (quando aplicável).
'   (2) Tenta forçar BackgroundQuery = False também nas conexões OLEDB (quando existirem).
'
' DEPENDÊNCIAS:
'   - Chamado por ThisWorkbook.Workbook_Open
'
' VERSÃO: 13/02/2026 (robustez + conexões + erro controlado)
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

' Desliga execução em 2º plano das consultas (sincroniza eventos de refresh)
Public Sub PQ_BackgroundOff()
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Dim lo As ListObject

    ' 1) Tabelas carregadas em planilha (ListObject -> QueryTable)
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

    ' 2) Conexões (quando houver OLEDBConnection disponível)
    ForceConnectionsBackgroundOff

    Exit Sub

TrataErro:
    ' Sem travar a abertura do arquivo: apenas avisa (se quiser, pode comentar o MsgBox)
    MsgBox "Erro em PQ_BackgroundOff: " & Err.Description, vbExclamation
End Sub

' Verifica se o ListObject tem QueryTable sem estourar erro
Private Function HasQueryTable(ByVal lo As ListObject) As Boolean
    On Error GoTo Falha
    HasQueryTable = Not (lo.QueryTable Is Nothing)
    Exit Function
Falha:
    HasQueryTable = False
End Function

' Força BackgroundQuery = False nas conexões que suportam OLEDBConnection
Private Sub ForceConnectionsBackgroundOff()
    On Error Resume Next

    Dim cn As WorkbookConnection
    For Each cn In ThisWorkbook.Connections
        If Not cn Is Nothing Then
            If Not cn.OLEDBConnection Is Nothing Then
                cn.OLEDBConnection.BackgroundQuery = False
            End If
        End If
    Next cn

    On Error GoTo 0
End Sub

' Hook "falso" para manter compatibilidade com chamadas existentes
Public Sub HookPQTables()
    ' Intencionalmente vazio – reservado para implementação futura se necessário.
End Sub

Public Sub ReHookPQ()
    HookPQTables
    MsgBox "ReHook executado (modo simplificado).", vbInformation
End Sub
