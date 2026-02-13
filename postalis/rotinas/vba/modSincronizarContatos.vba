Option Explicit

' =====================================================================================
' MÓDULO: modSincronizarContatos
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   (1) Sincronizar a tabela Registros_Contatos_Final com a Correção_Automática,
'       mantendo Status, Data Entrada/Saída e incluindo novos IDs.
'   (2) Sincronizar a tabela Registros_Contatos_Manual com a base oficial,
'       PRESERVANDO 100% DO HISTÓRICO MANUAL (NUNCA APAGAR LINHAS AUTOMATICAMENTE).
'
' REGRA DE LOG (IMPORTANTE):
'   - ESTE MÓDULO NÃO ESCREVE MAIS NENHUMA LINHA NA tblLogIDs.
'   - O LOG OFICIAL (Histórico dia-a-dia) é gerado SOMENTE no modExecucao, via Histórico Bitrix.
'
' VERSÃO: 12/02/2026
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

' Tabelas base
Private Const TBL_EXT    As String = "Correção_Automática"
Private Const TBL_CONT   As String = "Registros_Contatos_Final"

' Cabeçalhos padrão
Private Const COL_ID     As String = "ID"
Private Const COL_NOME   As String = "Nome"
Private Const COL_MAT    As String = "Matrícula"

' Colunas adicionais na tabela de contatos
Private Const COL_STATUS As String = "Status"
Private Const COL_DTENT  As String = "Data Entrada"
Private Const COL_DTSAI  As String = "Data Saída"

' Valores de status
Private Const ST_DENTRO  As String = "Dentro da extração"
Private Const ST_FORA    As String = "Fora da extração"

' Tabela de edição manual
Private Const SH_MANUAL  As String = "Registros Contatos Manual"
Private Const TBL_MANUAL As String = "Registros_Contatos_Manual"

' =====================================================================
' 1) SINCRONIZAÇÃO: REGISTROS_CONTATOS_FINAL x CORREÇÃO_AUTOMÁTICA (SEM LOG)
' =====================================================================
Public Sub Sincronizar_AteX_RegistrosContatos()
    On Error GoTo Tratar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim loExt As ListObject
    Dim loCont As ListObject
    Dim idxExt As Object
    Dim idxCont As Object
    Dim row As Long
    Dim idV As Variant
    Dim nomeV As Variant
    Dim matV As Variant
    Dim k As Variant
    Dim iRow As Long
    Dim curStatus As String
    Dim dtEnt As Variant
    Dim dtSai As Variant

    Set loExt = GetTableByName(TBL_EXT)
    Set loCont = GetTableByName(TBL_CONT)

    If loExt Is Nothing Or loCont Is Nothing Then
        MsgBox "Verifique as tabelas '" & TBL_EXT & "' e '" & TBL_CONT & "'.", vbExclamation
        GoTo Fim
    End If

    ' Garante colunas de trabalho no Registros Contatos
    EnsureListColumn loCont, COL_STATUS
    EnsureListColumn loCont, COL_DTENT
    EnsureListColumn loCont, COL_DTSAI

    ' Indexa IDs da extração
    Set idxExt = CreateObject("Scripting.Dictionary")
    idxExt.CompareMode = vbTextCompare

    If Not loExt.DataBodyRange Is Nothing Then
        For row = 1 To loExt.DataBodyRange.Rows.Count
            idV = ReadCell(loExt, row, COL_ID)
            If Len(Trim$(idV & "")) > 0 Then
                nomeV = ReadCell(loExt, row, COL_NOME)
                matV = ReadCell(loExt, row, COL_MAT)
                idxExt(CStr(idV)) = Array(nomeV, matV)
            End If
        Next row
    End If

    ' Indexa linhas existentes na tabela de contatos por ID
    Set idxCont = CreateObject("Scripting.Dictionary")
    idxCont.CompareMode = vbTextCompare

    If Not loCont.DataBodyRange Is Nothing Then
        For row = 1 To loCont.DataBodyRange.Rows.Count
            idV = ReadCell(loCont, row, COL_ID)
            If Len(Trim$(idV & "")) > 0 Then
                If Not idxCont.Exists(CStr(idV)) Then idxCont.Add CStr(idV), row
            End If
        Next row
    End If

    ' 1) Incluir IDs novos que estão na extração mas não existem em Registros Contatos
    For Each k In idxExt.Keys
        If Not idxCont.Exists(CStr(k)) Then
            Dim lr As ListRow
            Set lr = loCont.ListRows.Add
            iRow = lr.Index

            WriteCell loCont, iRow, COL_ID, k
            WriteCell loCont, iRow, COL_NOME, idxExt(k)(0)
            WriteCell loCont, iRow, COL_MAT, idxExt(k)(1)
            WriteCell loCont, iRow, COL_STATUS, ST_DENTRO
            WriteCell loCont, iRow, COL_DTENT, Now
            WriteCell loCont, iRow, COL_DTSAI, vbNullString
        End If
    Next k

    ' Reindexa depois de incluir novas linhas
    Set idxCont = CreateObject("Scripting.Dictionary")
    idxCont.CompareMode = vbTextCompare
    If Not loCont.DataBodyRange Is Nothing Then
        For row = 1 To loCont.DataBodyRange.Rows.Count
            idV = ReadCell(loCont, row, COL_ID)
            If Len(Trim$(idV & "")) > 0 Then
                If Not idxCont.Exists(CStr(idV)) Then idxCont.Add CStr(idV), row
            End If
        Next row
    End If

    ' 2) Atualizar Status, Data Entrada/Saída, Nome/Matrícula
    For Each k In idxCont.Keys
        iRow = idxCont(k)

        If idxExt.Exists(k) Then
            If Len(ReadCell(loCont, iRow, COL_NOME) & "") = 0 Then WriteCell loCont, iRow, COL_NOME, idxExt(k)(0)
            If Len(ReadCell(loCont, iRow, COL_MAT) & "") = 0 Then WriteCell loCont, iRow, COL_MAT, idxExt(k)(1)
        End If

        curStatus = CStr(ReadCell(loCont, iRow, COL_STATUS))
        dtEnt = ReadCell(loCont, iRow, COL_DTENT)
        dtSai = ReadCell(loCont, iRow, COL_DTSAI)

        If idxExt.Exists(k) Then
            If StrComp(curStatus, ST_DENTRO, vbTextCompare) <> 0 Then
                WriteCell loCont, iRow, COL_STATUS, ST_DENTRO
                If IsEmptyOrNull(dtEnt) Then WriteCell loCont, iRow, COL_DTENT, Now
                If Not IsEmptyOrNull(dtSai) Then WriteCell loCont, iRow, COL_DTSAI, vbNullString
            End If
        Else
            If StrComp(curStatus, ST_FORA, vbTextCompare) <> 0 Then
                WriteCell loCont, iRow, COL_STATUS, ST_FORA
                If IsEmptyOrNull(dtSai) Then WriteCell loCont, iRow, COL_DTSAI, Now
            End If
        End If
    Next k

Fim:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Tratar:
    MsgBox "Erro em Sincronizar_AteX_RegistrosContatos: " & Err.Description, vbExclamation
    Resume Fim
End Sub

' =====================================================================
' 2) SINCRONIZAR ABA REGISTROS CONTATOS MANUAL (NUNCA APAGA LINHAS)
' =====================================================================
Public Sub Sincronizar_RegistrosContatos_Manual()
    On Error GoTo Tratar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim loCont As ListObject
    Dim loMan  As ListObject
    Dim wsMan  As Worksheet
    Dim row As Long
    Dim idV As Variant
    Dim nomeV As Variant
    Dim matV As Variant
    Dim createdV As Variant
    Dim parentV As Variant

    Set loCont = GetTableByName(TBL_CONT)
    If loCont Is Nothing Then GoTo Fim
    If loCont.DataBodyRange Is Nothing Then GoTo Fim

    On Error Resume Next
    Set wsMan = ThisWorkbook.Worksheets(SH_MANUAL)
    On Error GoTo Tratar
    If wsMan Is Nothing Then GoTo Fim

    On Error Resume Next
    Set loMan = wsMan.ListObjects(TBL_MANUAL)
    On Error GoTo Tratar
    If loMan Is Nothing Then GoTo Fim

    EnsureListColumn loMan, COL_ID
    EnsureListColumn loMan, "Parent task ID"
    EnsureListColumn loMan, "Created on"
    EnsureListColumn loMan, COL_MAT
    EnsureListColumn loMan, COL_NOME

    Dim autoCols As Object
    Set autoCols = CreateObject("Scripting.Dictionary")
    autoCols.CompareMode = vbTextCompare
    autoCols.Add COL_ID, True
    autoCols.Add "Parent task ID", True
    autoCols.Add "Created on", True
    autoCols.Add COL_MAT, True
    autoCols.Add COL_NOME, True

    Dim manualCols As Collection: Set manualCols = New Collection
    Dim lc As ListColumn
    For Each lc In loMan.ListColumns
        If Not autoCols.Exists(lc.Name) Then manualCols.Add lc.Name
    Next lc

    Dim dictManual As Object: Set dictManual = CreateObject("Scripting.Dictionary")
    dictManual.CompareMode = vbTextCompare

    If Not loMan.DataBodyRange Is Nothing Then
        For row = 1 To loMan.DataBodyRange.Rows.Count
            idV = ReadCell(loMan, row, COL_ID)
            If Len(Trim$(idV & "")) > 0 Then
                For Each lc In loMan.ListColumns
                    If Not autoCols.Exists(lc.Name) Then
                        Dim key As String, val As Variant
                        key = CStr(idV) & "||" & lc.Name
                        val = ReadCell(loMan, row, lc.Name)
                        If Not IsEmptyOrNull(val) Then dictManual(key) = val
                    End If
                Next lc
            End If
        Next row
    End If

    Dim dictManRow As Object: Set dictManRow = CreateObject("Scripting.Dictionary")
    dictManRow.CompareMode = vbTextCompare

    If Not loMan.DataBodyRange Is Nothing Then
        For row = 1 To loMan.DataBodyRange.Rows.Count
            idV = ReadCell(loMan, row, COL_ID)
            If Len(Trim$(idV & "")) > 0 Then
                If Not dictManRow.Exists(CStr(idV)) Then dictManRow(CStr(idV)) = row
            End If
        Next row
    End If

    For row = 1 To loCont.DataBodyRange.Rows.Count
        idV = ReadCell(loCont, row, COL_ID)
        If Len(Trim$(idV & "")) = 0 Then GoTo ProximoID

        createdV = ReadCell(loCont, row, "Created on")
        nomeV = ReadCell(loCont, row, COL_NOME)
        matV = ReadCell(loCont, row, COL_MAT)
        parentV = ReadCell(loCont, row, "Parent task ID")

        Dim targetRow As Long
        If dictManRow.Exists(CStr(idV)) Then
            targetRow = CLng(dictManRow(CStr(idV)))
        Else
            Dim lr As ListRow
            Set lr = loMan.ListRows.Add
            targetRow = lr.Index
            dictManRow(CStr(idV)) = targetRow
        End If

        WriteCell loMan, targetRow, COL_ID, idV
        WriteCell loMan, targetRow, "Parent task ID", parentV
        WriteCell loMan, targetRow, "Created on", createdV
        WriteCell loMan, targetRow, COL_NOME, nomeV
        WriteCell loMan, targetRow, COL_MAT, matV

        Dim colName As Variant, keyMan As String
        For Each colName In manualCols
            keyMan = CStr(idV) & "||" & CStr(colName)
            If dictManual.Exists(keyMan) Then WriteCell loMan, targetRow, CStr(colName), dictManual(keyMan)
        Next colName

ProximoID:
    Next row

Fim:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Tratar:
    MsgBox "Erro em Sincronizar_RegistrosContatos_Manual: " & Err.Description, vbExclamation
    Resume Fim
End Sub

' =====================================================================
' HELPERS
' =====================================================================

Private Function GetTableByName(tblName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tblName, vbTextCompare) = 0 Then
                Set GetTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Sub EnsureListColumn(lo As ListObject, ByVal colName As String)
    Dim idx As Long
    On Error Resume Next
    idx = lo.ListColumns(colName).Index
    On Error GoTo 0
    If idx = 0 Then lo.ListColumns.Add.Name = colName
End Sub

Private Function ReadCell(lo As ListObject, r As Long, colName As String) As Variant
    On Error GoTo Falha
    ReadCell = lo.DataBodyRange.Cells(r, lo.ListColumns(colName).Index).Value
    Exit Function
Falha:
    ReadCell = vbNullString
End Function

Private Sub WriteCell(lo As ListObject, r As Long, colName As String, v As Variant)
    On Error Resume Next
    lo.DataBodyRange.Cells(r, lo.ListColumns(colName).Index).Value = v
End Sub

Private Function IsEmptyOrNull(v As Variant) As Boolean
    IsEmptyOrNull = (IsEmpty(v) Or IsNull(v) Or Len(Trim$(v & "")) = 0)
End Function
