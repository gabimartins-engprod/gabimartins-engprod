Option Explicit

' =====================================================================================
' MÓDULO: modExecucao
' SISTEMA: Rotina_Dados – Postalis
' BOTÃO ÚNICO: AtualizarTudo_E_Sincronizar
'
' REGRAS:
' - Log oficial: TABELA tblLogIDs (aba Log_Sincronizacao)
' - Backlog: ao clicar no botão, completa dias faltantes com base em tblHistoricoBitrix
' - Se houver mais de uma extração no dia, usa a ÚLTIMA (mais recente) via DataHoraExtracao
' - Ações exibidas: INICIO / (ENTROU/SAIU/RETORNOU se existirem) / RESUMO / FIM
' - ENTROU/SAIU/RETORNOU só aparecem se houver movimentação
' - Se NÃO houver extração em um dia: cria bloco com aviso (INICIO/RESUMO/FIM)
'
' PERÍODO MÍNIMO:
' - Processa desde START_DATE, e depois continua do "MARCO" gravado no log
'
' REGRA ESPECIAL (START_DATE):
' - No dia START_DATE (02/02/2026) NÃO EXISTE comparativo (sem dia anterior).
' - Portanto, nesse dia grava SOMENTE: INICIO + RESUMO (Total) + FIM.
'
' VERSÃO: 13/02/2026 (hotfix anti-duplicação + stamp por DataHoraExtracao + START_DATE sem comparativo)
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private isRunning As Boolean

Private Const SH_LOG As String = "Log_Sincronizacao"
Private Const TBL_LOG As String = "tblLogIDs"

Private Const SH_HIST As String = "Histórico Bitrix"
Private Const TB_HIST As String = "tblHistoricoBitrix"

Private Const COL_ID As String = "ID"
Private Const COL_DATAHORA_EXTRACAO As String = "DataHoraExtracao" ' <<< COLUNA OFICIAL (DateTime)

Private Const START_DATE As Date = #2/2/2026# ' 02/02/2026

Public Sub AtualizarTudo_E_Sincronizar()
    If isRunning Then Exit Sub
    isRunning = True
    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    Application.StatusBar = "Atualizando Power Query..."
    ThisWorkbook.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo TrataErro

    Application.StatusBar = "Processando sincronização interna..."
    Call Rotina_Sincronizar_E_Recalcular

    Application.StatusBar = "Atualizando Tabelas Dinâmicas..."
    RefreshAllPivots

    Application.StatusBar = "Gerando Log Histórico (dias faltantes)..."
    BuildLog_FromHistorico_FillMissingDays

Saida:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    isRunning = False
    Exit Sub

TrataErro:
    Application.StatusBar = "Erro: " & Err.Description
    Resume Saida
End Sub

' =====================================================================================
' LOG HISTÓRICO (COMPLETA DIAS FALTANTES) - COM TRAVA ANTI-DUPLICAÇÃO
' =====================================================================================

Private Sub BuildLog_FromHistorico_FillMissingDays()
    Dim loLog As ListObject: Set loLog = EnsureLogTable()
    Dim loHist As ListObject: Set loHist = FindTableByName(ThisWorkbook, TB_HIST)
    If loLog Is Nothing Then Exit Sub
    If loHist Is Nothing Then Exit Sub
    If loHist.DataBodyRange Is Nothing Then Exit Sub

    Dim idxID As Long: idxID = GetListColIndex(loHist, COL_ID)
    Dim idxDT As Long: idxDT = GetListColIndex(loHist, COL_DATAHORA_EXTRACAO)
    If idxID = 0 Or idxDT = 0 Then Exit Sub

    Dim maxHistDate As Date
    maxHistDate = GetMaxDateInHistorico(loHist)
    If maxHistDate = 0 Then Exit Sub
    If maxHistDate < START_DATE Then Exit Sub

    Dim lastProcessed As Date
    lastProcessed = GetLastProcessedDate(loLog)

    Dim firstToProcess As Date
    If lastProcessed = 0 Then
        firstToProcess = START_DATE
    Else
        firstToProcess = DateAdd("d", 1, lastProcessed)
    End If
    If firstToProcess > maxHistDate Then Exit Sub

    Dim dictLoggedDays As Object
    Set dictLoggedDays = BuildProcessedDaysDict(loLog) ' chave yyyy-mm-dd

    Dim dictDays As Object
    Set dictDays = BuildDaysWithExtractionDict(loHist, firstToProcess, maxHistDate)

    ' ? OPÇÃO A: SEM baseline anterior ao START_DATE
    Dim prevSet As Object: Set prevSet = CreateObject("Scripting.Dictionary"): prevSet.CompareMode = vbTextCompare
    Dim everSeen As Object: Set everSeen = CreateObject("Scripting.Dictionary"): everSeen.CompareMode = vbTextCompare

    Dim d As Date
    Dim lastAvailableDay As Date: lastAvailableDay = 0

    For d = firstToProcess To maxHistDate

        If DayAlreadyLogged(dictLoggedDays, d) Then GoTo ProximoDia

        Dim hasExtraction As Boolean
        hasExtraction = dictDays.Exists(Format(d, "yyyy-mm-dd"))

        Dim prevLabel As String
        If lastAvailableDay = 0 Then
            prevLabel = "(primeiro dia)"
        Else
            prevLabel = Format(lastAvailableDay, "dd/mm/yyyy")
        End If

        If Not hasExtraction Then
            Dim stampNo As Date
            stampNo = DateSerial(Year(d), Month(d), Day(d))

            AppendLog loLog, stampNo, "INICIO", "", "Comparativo Histórico Bitrix: " & prevLabel & " \ " & Format(d, "dd/mm/yyyy")
            AppendLog loLog, stampNo, "RESUMO", "", "DataArquivo: " & Format(d, "dd/mm/yyyy") & "; Não houve extração automática nessa data. Verificar o motivo com o Six."
            AppendLog loLog, stampNo, "FIM", "", "Comparativo concluído."

        Else
            Dim curSet As Object: Set curSet = CreateObject("Scripting.Dictionary"): curSet.CompareMode = vbTextCompare
            Dim maxDT As Date
            LoadIdsForLastExtractionOfDay loHist, d, curSet, maxDT
            If maxDT = 0 Then GoTo ProximoDia

            ' ? REGRA: START_DATE é o primeiro dia absoluto (sem comparativo)
            If d = START_DATE Then
                AppendLog loLog, maxDT, "INICIO", "", "Comparativo Histórico Bitrix: (primeiro dia) \ " & Format(d, "dd/mm/yyyy")
                AppendLog loLog, maxDT, "RESUMO", "", "DataArquivo: " & Format(d, "dd/mm/yyyy") & "; Total no dia: " & curSet.Count
                AppendLog loLog, maxDT, "FIM", "", "Comparativo concluído."

                prevSet.RemoveAll
                CopyDict curSet, prevSet
                MarkEverSeen curSet, everSeen
                lastAvailableDay = d

                dictLoggedDays(Format(d, "yyyy-mm-dd")) = True
                GoTo ProximoDia
            End If

            Dim entrou As Collection: Set entrou = New Collection
            Dim saiu As Collection: Set saiu = New Collection
            Dim retornou As Collection: Set retornou = New Collection
            Dim k As Variant

            For Each k In curSet.Keys
                If Not prevSet.Exists(CStr(k)) Then
                    If everSeen.Exists(CStr(k)) Then
                        retornou.Add CStr(k)
                    Else
                        entrou.Add CStr(k)
                        everSeen(CStr(k)) = True
                    End If
                End If
            Next k

            For Each k In prevSet.Keys
                If Not curSet.Exists(CStr(k)) Then saiu.Add CStr(k)
            Next k

            AppendLog loLog, maxDT, "INICIO", "", "Comparativo Histórico Bitrix: " & prevLabel & " \ " & Format(d, "dd/mm/yyyy")
            If retornou.Count > 0 Then WriteIds loLog, maxDT, "RETORNOU", retornou, "Voltou na extração (" & Format(d, "dd/mm/yyyy") & ")"
            If saiu.Count > 0 Then WriteIds loLog, maxDT, "SAIU", saiu, "Saiu da extração (" & Format(d, "dd/mm/yyyy") & ")"
            If entrou.Count > 0 Then WriteIds loLog, maxDT, "ENTROU", entrou, "Entrou na extração (" & Format(d, "dd/mm/yyyy") & ")"

            AppendLog loLog, maxDT, "RESUMO", "", _
                "DataArquivo: " & Format(d, "dd/mm/yyyy") & _
                "; ENTROU: " & entrou.Count & _
                "; RETORNOU: " & retornou.Count & _
                "; SAIU: " & saiu.Count & _
                "; Total no dia: " & curSet.Count

            AppendLog loLog, maxDT, "FIM", "", "Comparativo concluído."

            prevSet.RemoveAll
            CopyDict curSet, prevSet
            MarkEverSeen curSet, everSeen
            lastAvailableDay = d
        End If

        dictLoggedDays(Format(d, "yyyy-mm-dd")) = True

ProximoDia:
    Next d

    WriteMarco loLog, maxHistDate
End Sub

' =====================================================================================
' HISTÓRICO: ÚLTIMA EXTRAÇÃO DO DIA (mais recente) via DataHoraExtracao
' =====================================================================================

Private Sub LoadIdsForLastExtractionOfDay(ByVal loHist As ListObject, ByVal dayD As Date, ByRef dictOut As Object, ByRef maxDTOut As Date)
    dictOut.RemoveAll
    maxDTOut = 0
    If loHist.DataBodyRange Is Nothing Then Exit Sub

    Dim idxID As Long: idxID = GetListColIndex(loHist, COL_ID)
    Dim idxDT As Long: idxDT = GetListColIndex(loHist, COL_DATAHORA_EXTRACAO)
    If idxID = 0 Or idxDT = 0 Then Exit Sub

    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long
    Dim maxDT As Date: maxDT = 0

    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDT)) Then
            Dim dtV As Date
            dtV = CDate(arr(r, idxDT))
            If DateValue(dtV) = dayD Then
                If dtV > maxDT Then maxDT = dtV
            End If
        End If
    Next r

    If maxDT = 0 Then Exit Sub

    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDT)) And Not IsEmpty(arr(r, idxID)) Then
            Dim dt2 As Date
            dt2 = CDate(arr(r, idxDT))
            If dt2 = maxDT Then
                Dim idS As String
                idS = Trim$(CStr(arr(r, idxID)))
                If Len(idS) > 0 Then dictOut(idS) = True
            End If
        End If
    Next r

    maxDTOut = maxDT
End Sub

Private Function GetMaxDateInHistorico(ByVal loHist As ListObject) As Date
    Dim idxDT As Long: idxDT = GetListColIndex(loHist, COL_DATAHORA_EXTRACAO)
    If idxDT = 0 Then Exit Function

    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long, best As Date: best = 0
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDT)) Then
            Dim d As Date
            d = DateValue(CDate(arr(r, idxDT)))
            If d > best Then best = d
        End If
    Next r
    GetMaxDateInHistorico = best
End Function

Private Function BuildDaysWithExtractionDict(ByVal loHist As ListObject, ByVal d1 As Date, ByVal d2 As Date) As Object
    Dim idxDT As Long: idxDT = GetListColIndex(loHist, COL_DATAHORA_EXTRACAO)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If idxDT = 0 Then Set BuildDaysWithExtractionDict = dict: Exit Function
    If loHist.DataBodyRange Is Nothing Then Set BuildDaysWithExtractionDict = dict: Exit Function

    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDT)) Then
            Dim d As Date
            d = DateValue(CDate(arr(r, idxDT)))
            If d >= d1 And d <= d2 Then
                Dim key As String
                key = Format(d, "yyyy-mm-dd")
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        End If
    Next r

    Set BuildDaysWithExtractionDict = dict
End Function

' =====================================================================================
' LOG (tblLogIDs)
' =====================================================================================

Private Function EnsureLogTable() As ListObject
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lo = ws.ListObjects(TBL_LOG)
    On Error GoTo 0

    If lo Is Nothing Then
        ws.Range("A1:D1").Value = Array("DataHora", "Acao", "ID", "Obs")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D1"), , xlYes)
        lo.Name = TBL_LOG
        On Error Resume Next
        lo.TableStyle = "TableStyleMedium2"
        On Error GoTo 0
        ws.Columns("A:D").EntireColumn.AutoFit
    End If

    Set EnsureLogTable = lo
End Function

Private Sub AppendLog(ByVal loLog As ListObject, ByVal dt As Date, ByVal acao As String, ByVal idVal As String, ByVal obs As String)
    Dim lr As ListRow
    Set lr = loLog.ListRows.Add
    lr.Range.Cells(1, 1).Value = dt
    lr.Range.Cells(1, 2).Value = acao
    lr.Range.Cells(1, 3).Value = idVal
    lr.Range.Cells(1, 4).Value = obs
End Sub

Private Sub WriteIds(ByVal loLog As ListObject, ByVal dt As Date, ByVal acao As String, ByVal col As Collection, ByVal obs As String)
    Dim i As Long
    For i = 1 To col.Count
        AppendLog loLog, dt, acao, CStr(col(i)), obs
    Next i
End Sub

Private Function GetLastProcessedDate(ByVal loLog As ListObject) As Date
    GetLastProcessedDate = 0
    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function

    Dim arr As Variant: arr = loLog.DataBodyRange.Value
    Dim r As Long
    For r = UBound(arr, 1) To 1 Step -1
        If UCase$(Trim$(CStr(arr(r, 2)))) = "MARCO" Then
            GetLastProcessedDate = ParseMarcoDate(CStr(arr(r, 4)))
            Exit Function
        End If
    Next r
End Function

Private Sub WriteMarco(ByVal loLog As ListObject, ByVal d As Date)
    AppendLog loLog, Now, "MARCO", "", "UltimaDataArquivoProcessada=" & Format(d, "yyyy-mm-dd")
End Sub

Private Function ParseMarcoDate(ByVal s As String) As Date
    On Error GoTo Fim
    Dim p As Long: p = InStr(1, s, "=", vbTextCompare)
    If p = 0 Then GoTo Fim
    Dim v As String: v = Trim$(mid$(s, p + 1))
    If Len(v) >= 10 Then
        ParseMarcoDate = DateSerial(CInt(mid$(v, 1, 4)), CInt(mid$(v, 6, 2)), CInt(mid$(v, 9, 2)))
        Exit Function
    End If
Fim:
    ParseMarcoDate = 0
End Function

Private Function BuildProcessedDaysDict(ByVal loLog As ListObject) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set BuildProcessedDaysDict = dict

    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function

    Dim arr As Variant: arr = loLog.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, 2)))) = "RESUMO" Then
            Dim d As Date
            d = ParseDataArquivoFromObs(CStr(arr(r, 4)))
            If d > 0 Then dict(Format(d, "yyyy-mm-dd")) = True
        End If
    Next r
End Function

Private Function DayAlreadyLogged(ByVal dictLoggedDays As Object, ByVal d As Date) As Boolean
    DayAlreadyLogged = False
    If dictLoggedDays Is Nothing Then Exit Function
    DayAlreadyLogged = dictLoggedDays.Exists(Format(d, "yyyy-mm-dd"))
End Function

Private Function ParseDataArquivoFromObs(ByVal obs As String) As Date
    On Error GoTo Fim
    Dim p As Long: p = InStr(1, obs, "DataArquivo:", vbTextCompare)
    If p = 0 Then GoTo Fim

    Dim s As String: s = Trim$(mid$(obs, p + Len("DataArquivo:")))
    If Len(s) < 10 Then GoTo Fim

    Dim dd As Integer, mm As Integer, yy As Integer
    dd = CInt(mid$(s, 1, 2))
    mm = CInt(mid$(s, 4, 2))
    yy = CInt(mid$(s, 7, 4))

    ParseDataArquivoFromObs = DateSerial(yy, mm, dd)
    Exit Function
Fim:
    ParseDataArquivoFromObs = 0
End Function

' =====================================================================================
' HELPERS
' =====================================================================================

Private Function FindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindTableByName = Nothing
End Function

Private Function GetListColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), colName, vbTextCompare) = 0 Then
            GetListColIndex = lc.Index
            Exit Function
        End If
    Next lc
    GetListColIndex = 0
End Function

Private Sub CopyDict(ByVal src As Object, ByRef dst As Object)
    dst.RemoveAll
    Dim k As Variant
    For Each k In src.Keys
        dst(CStr(k)) = True
    Next k
End Sub

Private Sub MarkEverSeen(ByVal src As Object, ByRef everSeen As Object)
    Dim k As Variant
    For Each k In src.Keys
        If Not everSeen.Exists(CStr(k)) Then everSeen(CStr(k)) = True
    Next k
End Sub

Private Sub RefreshAllPivots()
    On Error Resume Next
    Dim ws As Worksheet, pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    On Error GoTo 0
End Sub
