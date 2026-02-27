Option Explicit

' =====================================================================================
' MÓDULO: modSincronizarContatos
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   (1) Sincronizar Registros_Contatos_Final com Correção_Automática:
'       - incluir novos IDs da extração
'       - atualizar Nome/Matrícula quando disponível
'       - manter Status e preencher Datas (Entrada/Saída) conforme presença na extração
'   (2) Sincronizar Registros_Contatos_Manual:
'       - incluir IDs novos (a partir do Final e/ou do Histórico)
'       - backfill leve de datas (Entrada/Saída/Retorno) apenas para pendentes
'
' ORIGEM DOS DADOS:
'   - Tabela "Correção_Automática" (TBL_EXT)
'   - Tabela "Registros_Contatos_Final" (TBL_CONT)
'   - Tabela "tblHistoricoBitrix" (TB_HIST) na aba "Histórico Bitrix"
'
' DESTINO:
'   - Tabela "Registros_Contatos_Final" (atualização de Status/Data Entrada/Data Saída e inclusão de IDs)
'   - Tabela "Registros_Contatos_Manual" (inclusão de IDs e preenchimento de datas pendentes)
'
' COMPORTAMENTO TÉCNICO:
'   - BASELINE conceitual: 01/02/2026.
'   - Backfill/ciclo só percorre DIAS COM EXTRAÇÃO (DataArquivo no Histórico).
'   - DataHoraExtracao (quando existir) é usada para definir "foto válida" do dia (última extração).
'
' MACROS PÚBLICAS:
'   - Sincronizar_AteX_RegistrosContatos: sincroniza Final com a Extração (Dentro/Fora + datas).
'   - Manual_IncluirNovosIDs_PeloFinal: inclui IDs novos na Manual a partir do Final.
'   - Manual_IncluirIDs_PeloHistorico_0202: inclui IDs novos direto do Histórico (linha mais recente por ID).
'   - Manual_Backfill_Rapido_0202: wrapper para backfill leve (pendentes) na Manual.
'
' DEPENDÊNCIAS:
'   - Chamado por: modExecucao.AtualizarTudo_E_Sincronizar
'
' VERSÃO: 25/02/2026 (baseline 01/02 + mantém regra "dias com extração")
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private Const TBL_EXT As String = "Correção_Automática"
Private Const TBL_CONT As String = "Registros_Contatos_Final"
Private Const SH_MANUAL As String = "Registros Contatos Manual"
Private Const TBL_MANUAL As String = "Registros_Contatos_Manual"

' Histórico Bitrix
Private Const TB_HIST As String = "tblHistoricoBitrix"
Private Const COL_DATAARQUIVO As String = "DataArquivo"               ' Date
Private Const COL_DATAHORA_EXTRACAO As String = "DataHoraExtracao"    ' DateTime (opcional)

Private Const COL_ID As String = "ID"
Private Const COL_NOME As String = "Nome"
Private Const COL_MAT As String = "Matrícula"

Private Const COL_STATUS As String = "Status"
Private Const COL_DTENT As String = "Data Entrada"
Private Const COL_DTSAI As String = "Data Saída"

Private Const ST_DENTRO As String = "Dentro da extração"
Private Const ST_FORA As String = "Fora da extração"

Private Const BASELINE_D As Date = #2/1/2026# ' 01/02/2026

' ===========================================================
' 1) SINCRONIZAÇÃO PRINCIPAL (FINAL)
' ===========================================================

Public Sub Sincronizar_AteX_RegistrosContatos()

    On Error GoTo Tratar

    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim loExt As ListObject, loCont As ListObject
    Set loExt = GetTableByName(TBL_EXT)
    Set loCont = GetTableByName(TBL_CONT)

    If loExt Is Nothing Then MsgBox "Tabela '" & TBL_EXT & "' não encontrada.", vbExclamation: GoTo Fim
    If loCont Is Nothing Then MsgBox "Tabela '" & TBL_CONT & "' não encontrada.", vbExclamation: GoTo Fim

    Dim dataRef As Date
    dataRef = GetLatestHistoricoDateOrToday()
    If dataRef < BASELINE_D Then dataRef = BASELINE_D

    Dim idxExtID As Long, idxExtNome As Long, idxExtMat As Long
    idxExtID = GetListColIndex(loExt, COL_ID)
    idxExtNome = GetListColIndex(loExt, COL_NOME)
    idxExtMat = GetListColIndex(loExt, COL_MAT)

    If idxExtID = 0 Then
        MsgBox "Coluna '" & COL_ID & "' não encontrada na '" & TBL_EXT & "'.", vbExclamation
        GoTo Fim
    End If

    Dim idxContID As Long, idxContNome As Long, idxContMat As Long
    idxContID = GetListColIndex(loCont, COL_ID)
    idxContNome = GetListColIndex(loCont, COL_NOME)
    idxContMat = GetListColIndex(loCont, COL_MAT)

    If idxContID = 0 Then MsgBox "Coluna '" & COL_ID & "' não encontrada na '" & TBL_CONT & "'.", vbExclamation: GoTo Fim
    If idxContNome = 0 Then MsgBox "Coluna '" & COL_NOME & "' não encontrada na '" & TBL_CONT & "'.", vbExclamation: GoTo Fim
    If idxContMat = 0 Then MsgBox "Coluna '" & COL_MAT & "' não encontrada na '" & TBL_CONT & "'.", vbExclamation: GoTo Fim

    If GetListColIndex(loCont, "Created on") = 0 Then MsgBox "Coluna 'Created on' não encontrada na '" & TBL_CONT & "'.", vbExclamation: GoTo Fim
    If GetListColIndex(loCont, "Parent task ID") = 0 Then MsgBox "Coluna 'Parent task ID' não encontrada na '" & TBL_CONT & "'.", vbExclamation: GoTo Fim

    EnsureListColumn loCont, COL_STATUS
    EnsureListColumn loCont, COL_DTENT
    EnsureListColumn loCont, COL_DTSAI

    Dim idxContStatus As Long, idxContEnt As Long, idxContSai As Long
    idxContStatus = GetListColIndex(loCont, COL_STATUS)
    idxContEnt = GetListColIndex(loCont, COL_DTENT)
    idxContSai = GetListColIndex(loCont, COL_DTSAI)

    ApplyTextFormatColumn loExt, COL_ID
    ApplyTextFormatColumn loCont, COL_ID

    Dim dictExt As Object: Set dictExt = CreateObject("Scripting.Dictionary")
    dictExt.CompareMode = vbTextCompare

    Dim row As Long
    If Not loExt.DataBodyRange Is Nothing Then
        For row = 1 To loExt.DataBodyRange.Rows.Count
            Dim idExt As String
            idExt = Trim$(CStr(loExt.DataBodyRange.Cells(row, idxExtID).Value))
            If Len(idExt) > 0 Then
                Dim nomeV As Variant, matV As Variant
                nomeV = vbNullString
                matV = vbNullString
                If idxExtNome > 0 Then nomeV = loExt.DataBodyRange.Cells(row, idxExtNome).Value
                If idxExtMat > 0 Then matV = loExt.DataBodyRange.Cells(row, idxExtMat).Value
                dictExt(idExt) = Array(nomeV, matV)
            End If
        Next row
    End If

    Dim dictRow As Object: Set dictRow = CreateObject("Scripting.Dictionary")
    dictRow.CompareMode = vbTextCompare

    If Not loCont.DataBodyRange Is Nothing Then
        For row = 1 To loCont.DataBodyRange.Rows.Count
            Dim idCont As String
            idCont = Trim$(CStr(loCont.DataBodyRange.Cells(row, idxContID).Value))
            If Len(idCont) > 0 Then
                If Not dictRow.Exists(idCont) Then dictRow.Add idCont, row
            End If
        Next row
    End If

    Dim k As Variant
    For Each k In dictExt.keys
        If Not dictRow.Exists(CStr(k)) Then
            Dim lr As ListRow
            Set lr = loCont.ListRows.Add

            lr.Range.Cells(1, idxContID).Value = CStr(k)

            If Len(Trim$(CStr(dictExt(CStr(k))(0)))) > 0 Then lr.Range.Cells(1, idxContNome).Value = dictExt(CStr(k))(0)
            If Len(Trim$(CStr(dictExt(CStr(k))(1)))) > 0 Then lr.Range.Cells(1, idxContMat).Value = dictExt(CStr(k))(1)

            If idxContStatus > 0 Then lr.Range.Cells(1, idxContStatus).Value = ST_DENTRO
            If idxContEnt > 0 Then lr.Range.Cells(1, idxContEnt).Value = dataRef
            If idxContSai > 0 Then lr.Range.Cells(1, idxContSai).ClearContents

            If Not loCont.DataBodyRange Is Nothing Then
                dictRow(CStr(k)) = loCont.DataBodyRange.Rows.Count
            End If
        End If
    Next k

    If Not loCont.DataBodyRange Is Nothing Then
        For row = 1 To loCont.DataBodyRange.Rows.Count

            Dim idNow As String
            idNow = Trim$(CStr(loCont.DataBodyRange.Cells(row, idxContID).Value))
            If Len(idNow) = 0 Then GoTo ProxLinha

            Dim estaNaExtracao As Boolean
            estaNaExtracao = dictExt.Exists(idNow)

            Dim statusAtual As String
            statusAtual = ""
            If idxContStatus > 0 Then statusAtual = Trim$(CStr(loCont.DataBodyRange.Cells(row, idxContStatus).Value))

            If estaNaExtracao Then
                If idxContStatus > 0 Then loCont.DataBodyRange.Cells(row, idxContStatus).Value = ST_DENTRO

                If Len(Trim$(CStr(dictExt(idNow)(0)))) > 0 Then loCont.DataBodyRange.Cells(row, idxContNome).Value = dictExt(idNow)(0)
                If Len(Trim$(CStr(dictExt(idNow)(1)))) > 0 Then loCont.DataBodyRange.Cells(row, idxContMat).Value = dictExt(idNow)(1)

                If idxContSai > 0 Then
                    If UCase$(statusAtual) = UCase$(ST_FORA) Then loCont.DataBodyRange.Cells(row, idxContSai).ClearContents
                End If

                If idxContEnt > 0 Then
                    If IsEmptyOrBlank(loCont.DataBodyRange.Cells(row, idxContEnt).Value) Then loCont.DataBodyRange.Cells(row, idxContEnt).Value = dataRef
                End If
            Else
                If idxContStatus > 0 Then loCont.DataBodyRange.Cells(row, idxContStatus).Value = ST_FORA
                If idxContSai > 0 Then
                    If IsEmptyOrBlank(loCont.DataBodyRange.Cells(row, idxContSai).Value) Then loCont.DataBodyRange.Cells(row, idxContSai).Value = dataRef
                End If
            End If

ProxLinha:
        Next row
    End If

Fim:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Tratar:
    MsgBox "Erro em Sincronizar_AteX_RegistrosContatos: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Fim

End Sub

' =====================================================================================
' WRAPPER PARA modExecucao
' =====================================================================================
Public Sub Manual_Backfill_Rapido_0202()
    Manual_Backfill_Pendentes_0202
End Sub

' =====================================================================================
' BACKFILL LEVE (SOMENTE PENDENTES) — ciclo só em dias com extração
' BASELINE conceitual: 01/02/2026 (mas a varredura usa apenas dias com extração)
' =====================================================================================
Public Sub Manual_Backfill_Pendentes_0202()
    On Error GoTo Trata

    Dim prevCalc As XlCalculation, prevEvents As Boolean, prevScreen As Boolean
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim baseMin As Date
    baseMin = BASELINE_D

    ' =========================
    ' Tabela MANUAL
    ' =========================
    Dim wsMan As Worksheet, loMan As ListObject
    Set wsMan = ThisWorkbook.Worksheets(SH_MANUAL)
    Set loMan = wsMan.ListObjects(TBL_MANUAL)
    If loMan Is Nothing Then GoTo Fim
    If loMan.DataBodyRange Is Nothing Then GoTo Fim

    EnsureListColumn loMan, "Data Entrada"
    EnsureListColumn loMan, "Data Saída"
    EnsureListColumn loMan, "Data Retorno"

    Dim idxManID As Long, idxEnt As Long, idxSai As Long, idxRet As Long
    idxManID = GetListColIndex(loMan, COL_ID)
    idxEnt = GetListColIndex(loMan, "Data Entrada")
    idxSai = GetListColIndex(loMan, "Data Saída")
    idxRet = GetListColIndex(loMan, "Data Retorno")
    If idxManID = 0 Or idxEnt = 0 Or idxSai = 0 Or idxRet = 0 Then GoTo Fim

    ApplyTextFormatColumn loMan, COL_ID

    ' =========================
    ' Tabela HISTÓRICO
    ' =========================
    Dim loHist As ListObject
    Set loHist = GetTableByName(TB_HIST)
    If loHist Is Nothing Then GoTo Fim
    If loHist.DataBodyRange Is Nothing Then GoTo Fim

    Dim idxH_ID As Long, idxH_DA As Long, idxH_DT As Long
    idxH_ID = GetListColIndex(loHist, COL_ID)
    idxH_DA = GetListColIndex(loHist, COL_DATAARQUIVO)
    idxH_DT = GetListColIndex(loHist, COL_DATAHORA_EXTRACAO)

    If idxH_ID = 0 Then GoTo Fim
    If idxH_DA = 0 And idxH_DT = 0 Then GoTo Fim

    ApplyTextFormatColumn loHist, COL_ID

    Dim arr As Variant
    arr = loHist.DataBodyRange.Value

    ' dMax do histórico
    Dim dMax As Date: dMax = 0
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim dRow As Date: dRow = 0
        If idxH_DA > 0 Then
            dRow = Hist_ToDateOnly(arr(r, idxH_DA))
        ElseIf idxH_DT > 0 Then
            dRow = Hist_ToDateOnly(arr(r, idxH_DT))
        End If
        If dRow >= baseMin Then
            If dMax = 0 Or dRow > dMax Then dMax = dRow
        End If
    Next r
    If dMax = 0 Then GoTo Fim

    ' Dias com extração (DataArquivo distintos), ordenados
    Dim extrDates As Variant
    extrDates = GetExtractionDatesSorted(arr, idxH_DA, baseMin, dMax)
    If IsEmpty(extrDates) Then GoTo Fim

    ' dictDays: dias com extração (para validar datas na Manual)
    Dim dictDays As Object
    Set dictDays = BuildExtractionDaysDict(arr, idxH_DA, baseMin, dMax)

    ' =========================
    ' 1) Monta PENDENTES (inclui também datas inválidas)
    ' =========================
    Dim pend As Object: Set pend = CreateObject("Scripting.Dictionary")
    pend.CompareMode = vbTextCompare

    Dim rr As Long, idm As String
    For rr = 1 To loMan.DataBodyRange.Rows.Count
        idm = Trim$(CStr(loMan.DataBodyRange.Cells(rr, idxManID).Value))
        If Len(idm) > 0 Then

            Dim vEnt As Variant, vSai As Variant, vRet As Variant
            vEnt = loMan.DataBodyRange.Cells(rr, idxEnt).Value
            vSai = loMan.DataBodyRange.Cells(rr, idxSai).Value
            vRet = loMan.DataBodyRange.Cells(rr, idxRet).Value

            If Not IsEmptyOrBlank(vSai) Then
                If Not IsExtractionDay(dictDays, vSai) Then loMan.DataBodyRange.Cells(rr, idxSai).ClearContents
            End If
            If Not IsEmptyOrBlank(vRet) Then
                If Not IsExtractionDay(dictDays, vRet) Then loMan.DataBodyRange.Cells(rr, idxRet).ClearContents
            End If

            If IsEmptyOrBlank(loMan.DataBodyRange.Cells(rr, idxSai).Value) Then
                loMan.DataBodyRange.Cells(rr, idxRet).ClearContents
            End If

            vEnt = loMan.DataBodyRange.Cells(rr, idxEnt).Value
            vSai = loMan.DataBodyRange.Cells(rr, idxSai).Value
            vRet = loMan.DataBodyRange.Cells(rr, idxRet).Value

            If IsEmptyOrBlank(vEnt) Or IsEmptyOrBlank(vSai) Or IsEmptyOrBlank(vRet) Then
                pend(idm) = rr
            End If
        End If
    Next rr
    If pend.Count = 0 Then GoTo Fim

    ' =========================
    ' 2) Última extração do dia (se DataHoraExtracao existir)
    ' =========================
    Dim dayMaxDT As Object: Set dayMaxDT = CreateObject("Scripting.Dictionary")
    dayMaxDT.CompareMode = vbTextCompare

    If idxH_DT > 0 And idxH_DA > 0 Then
        For r = 1 To UBound(arr, 1)
            If IsEmpty(arr(r, idxH_ID)) Then GoTo P1Next
            If IsEmpty(arr(r, idxH_DA)) Then GoTo P1Next
            If IsEmpty(arr(r, idxH_DT)) Or Not IsDate(arr(r, idxH_DT)) Then GoTo P1Next

            Dim dA As Date: dA = Hist_ToDateOnly(arr(r, idxH_DA))
            If dA = 0 Or dA < baseMin Or dA > dMax Then GoTo P1Next

            Dim key As String: key = Format(dA, "yyyy-mm-dd")
            Dim dtV As Date: dtV = CDate(arr(r, idxH_DT))
            If Not dayMaxDT.Exists(key) Then
                dayMaxDT(key) = dtV
            Else
                If dtV > CDate(dayMaxDT(key)) Then dayMaxDT(key) = dtV
            End If
P1Next:
        Next r
    End If

    ' =========================
    ' 3) presence(ID)(dayKey)=True e firstSeen(ID)=min day (somente "foto válida" do dia)
    ' =========================
    Dim presence As Object: Set presence = CreateObject("Scripting.Dictionary")
    presence.CompareMode = vbTextCompare

    Dim firstSeen As Object: Set firstSeen = CreateObject("Scripting.Dictionary")
    firstSeen.CompareMode = vbTextCompare

    Dim idH As String, dA2 As Date, dayKey As String, includeRow As Boolean
    Dim dctPres As Object

    For r = 1 To UBound(arr, 1)

        If IsEmpty(arr(r, idxH_ID)) Then GoTo P2Next
        If idxH_DA > 0 Then
            If IsEmpty(arr(r, idxH_DA)) Then GoTo P2Next
        Else
            GoTo P2Next
        End If

        idH = Trim$(CStr(arr(r, idxH_ID)))
        If Len(idH) = 0 Then GoTo P2Next
        If Not pend.Exists(idH) Then GoTo P2Next

        dA2 = Hist_ToDateOnly(arr(r, idxH_DA))
        If dA2 = 0 Or dA2 < baseMin Or dA2 > dMax Then GoTo P2Next

        dayKey = Format(dA2, "yyyy-mm-dd")

        includeRow = True
        If idxH_DT > 0 And dayMaxDT.Exists(dayKey) Then
            If IsEmpty(arr(r, idxH_DT)) Or Not IsDate(arr(r, idxH_DT)) Then
                includeRow = False
            Else
                includeRow = (CDate(arr(r, idxH_DT)) = CDate(dayMaxDT(dayKey)))
            End If
        End If
        If Not includeRow Then GoTo P2Next

        If Not firstSeen.Exists(idH) Then
            firstSeen(idH) = dA2
        Else
            If dA2 < CDate(firstSeen(idH)) Then firstSeen(idH) = dA2
        End If

        If Not presence.Exists(idH) Then
            Set dctPres = CreateObject("Scripting.Dictionary")
            dctPres.CompareMode = vbTextCompare
            presence.Add idH, dctPres
        Else
            Set dctPres = presence(idH)
        End If
        dctPres(dayKey) = True

P2Next:
    Next r

    ' =========================
    ' 4) Calcula ciclo por ID — só em DIAS COM EXTRAÇÃO (extrDates)
    ' =========================
    Dim idK As Variant
    Dim i As Long
    For Each idK In pend.keys

        rr = CLng(pend(CStr(idK)))

        Dim dFirst As Date
        If firstSeen.Exists(CStr(idK)) Then
            dFirst = CDate(firstSeen(CStr(idK)))
        Else
            GoTo NextId
        End If

        Dim lastExit As Date: lastExit = 0
        Dim lastReturn As Date: lastReturn = 0
        Dim hadExit As Boolean: hadExit = False

        Dim inPrev As Boolean: inPrev = False
        Dim inNow As Boolean

        Dim dctOne As Object
        If presence.Exists(CStr(idK)) Then
            Set dctOne = presence(CStr(idK))
        Else
            Set dctOne = Nothing
        End If

        For i = LBound(extrDates) To UBound(extrDates)

            Dim d As Date: d = CDate(extrDates(i))
            If d < dFirst Then GoTo NextDay

            dayKey = Format(d, "yyyy-mm-dd")

            inNow = False
            If Not dctOne Is Nothing Then
                If dctOne.Exists(dayKey) Then inNow = True
            End If

            If inPrev And (Not inNow) Then
                lastExit = d
                hadExit = True
            End If

            If (Not inPrev) And inNow Then
                If hadExit Then lastReturn = d
            End If

            inPrev = inNow

NextDay:
        Next i

        If Len(Trim$(CStr(loMan.DataBodyRange.Cells(rr, idxEnt).Value))) = 0 Then
            loMan.DataBodyRange.Cells(rr, idxEnt).Value = dFirst
        End If
        If lastExit <> 0 Then loMan.DataBodyRange.Cells(rr, idxSai).Value = lastExit
        If lastReturn <> 0 Then loMan.DataBodyRange.Cells(rr, idxRet).Value = lastReturn

NextId:
    Next idK

    MsgBox "DIAG Backfill Pendentes:" & vbCrLf & _
           "Pendentes (Manual) = " & pend.Count & vbCrLf & _
           "firstSeen (Histórico) = " & firstSeen.Count & vbCrLf & _
           "dMax (Histórico) = " & Format(dMax, "dd/mm/yyyy"), vbInformation

Fim:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Trata:
    MsgBox "Erro em Manual_Backfill_Pendentes_0202: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Fim
End Sub

' ===========================================================
' INCLUI NOVOS IDS NA MANUAL A PARTIR DO FINAL
' ===========================================================
Public Sub Manual_IncluirNovosIDs_PeloFinal()
    On Error GoTo Trata

    Dim loFinal As ListObject, wsMan As Worksheet, loMan As ListObject
    Set loFinal = GetTableByName(TBL_CONT)
    If loFinal Is Nothing Then Exit Sub
    If loFinal.DataBodyRange Is Nothing Then Exit Sub

    Set wsMan = ThisWorkbook.Worksheets(SH_MANUAL)
    Set loMan = wsMan.ListObjects(TBL_MANUAL)
    If loMan Is Nothing Then Exit Sub

    ApplyTextFormatColumn loFinal, COL_ID
    ApplyTextFormatColumn loMan, COL_ID

    Dim idxFID As Long, idxFParent As Long, idxFCreated As Long, idxFMat As Long, idxFNome As Long
    idxFID = GetListColIndex(loFinal, COL_ID)
    idxFParent = GetListColIndex(loFinal, "Parent task ID")
    idxFCreated = GetListColIndex(loFinal, "Created on")
    idxFMat = GetListColIndex(loFinal, COL_MAT)
    idxFNome = GetListColIndex(loFinal, COL_NOME)

    If idxFID = 0 Or idxFParent = 0 Or idxFCreated = 0 Or idxFMat = 0 Or idxFNome = 0 Then Exit Sub

    Dim idxMID As Long, idxMParent As Long, idxMCreated As Long, idxMMat As Long, idxMNome As Long
    idxMID = GetListColIndex(loMan, COL_ID)
    idxMParent = GetListColIndex(loMan, "Parent task ID")
    idxMCreated = GetListColIndex(loMan, "Created on")
    idxMMat = GetListColIndex(loMan, COL_MAT)
    idxMNome = GetListColIndex(loMan, COL_NOME)

    If idxMID = 0 Or idxMParent = 0 Or idxMCreated = 0 Or idxMMat = 0 Or idxMNome = 0 Then Exit Sub

    Dim dictMan As Object: Set dictMan = CreateObject("Scripting.Dictionary")
    dictMan.CompareMode = vbTextCompare

    Dim r As Long
    If Not loMan.DataBodyRange Is Nothing Then
        For r = 1 To loMan.DataBodyRange.Rows.Count
            Dim idm As String
            idm = Trim$(CStr(loMan.DataBodyRange.Cells(r, idxMID).Value))
            If Len(idm) > 0 Then dictMan(idm) = True
        Next r
    End If

    For r = 1 To loFinal.DataBodyRange.Rows.Count
        Dim idf As String
        idf = Trim$(CStr(loFinal.DataBodyRange.Cells(r, idxFID).Value))
        If Len(idf) = 0 Then GoTo Prox

        If Not dictMan.Exists(idf) Then
            Dim lr As ListRow
            Set lr = loMan.ListRows.Add
            lr.Range.Cells(1, idxMID).Value = idf
            lr.Range.Cells(1, idxMParent).Value = loFinal.DataBodyRange.Cells(r, idxFParent).Value
            lr.Range.Cells(1, idxMCreated).Value = loFinal.DataBodyRange.Cells(r, idxFCreated).Value
            lr.Range.Cells(1, idxMMat).Value = loFinal.DataBodyRange.Cells(r, idxFMat).Value
            lr.Range.Cells(1, idxMNome).Value = loFinal.DataBodyRange.Cells(r, idxFNome).Value
            dictMan(idf) = True
        End If
Prox:
    Next r

    Exit Sub
Trata:
    MsgBox "Erro em Manual_IncluirNovosIDs_PeloFinal: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ===========================================================
' INCLUI IDs NA MANUAL DIRETO DO HISTÓRICO (desde 01/02/2026)
' Critério: pega a linha MAIS RECENTE do ID (DataArquivo maior; desempate por DataHoraExtracao)
' ===========================================================
Public Sub Manual_IncluirIDs_PeloHistorico_0202()
    On Error GoTo Trata

    Dim baseMin As Date: baseMin = BASELINE_D

    Dim wsMan As Worksheet, loMan As ListObject
    Set wsMan = ThisWorkbook.Worksheets(SH_MANUAL)
    Set loMan = wsMan.ListObjects(TBL_MANUAL)
    If loMan Is Nothing Then Exit Sub

    Dim loHist As ListObject
    Set loHist = GetTableByName(TB_HIST)
    If loHist Is Nothing Then Exit Sub
    If loHist.DataBodyRange Is Nothing Then Exit Sub

    Dim idxMID As Long: idxMID = GetListColIndex(loMan, COL_ID)
    If idxMID = 0 Then Exit Sub

    Dim idxMParent As Long: idxMParent = GetListColIndexFlex(loMan, "Parent task ID")
    Dim idxMCreated As Long: idxMCreated = GetListColIndexFlex(loMan, "Created on")
    Dim idxMMat As Long: idxMMat = GetListColIndexFlex(loMan, COL_MAT)
    Dim idxMNome As Long: idxMNome = GetListColIndexFlex(loMan, COL_NOME)

    Dim idxHID As Long: idxHID = GetListColIndex(loHist, COL_ID)
    Dim idxHDA As Long: idxHDA = GetListColIndex(loHist, COL_DATAARQUIVO)
    If idxHID = 0 Or idxHDA = 0 Then Exit Sub

    Dim idxHDT As Long: idxHDT = GetListColIndexFlex(loHist, COL_DATAHORA_EXTRACAO)
    Dim idxHParent As Long: idxHParent = GetListColIndexFlex(loHist, "Parent task ID")
    Dim idxHCreated As Long: idxHCreated = GetListColIndexFlex(loHist, "Created on")
    Dim idxHMat As Long: idxHMat = GetListColIndexFlex(loHist, COL_MAT)
    Dim idxHNome As Long: idxHNome = GetListColIndexFlex(loHist, COL_NOME)

    ApplyTextFormatColumn loMan, COL_ID
    ApplyTextFormatColumn loHist, COL_ID

    Dim dictMan As Object: Set dictMan = CreateObject("Scripting.Dictionary")
    dictMan.CompareMode = vbTextCompare

    Dim r As Long, idm As String
    If Not loMan.DataBodyRange Is Nothing Then
        For r = 1 To loMan.DataBodyRange.Rows.Count
            idm = Trim$(CStr(loMan.DataBodyRange.Cells(r, idxMID).Value))
            If Len(idm) > 0 Then dictMan(idm) = True
        Next r
    End If

    Dim arr As Variant: arr = loHist.DataBodyRange.Value

    Dim bestRow As Object: Set bestRow = CreateObject("Scripting.Dictionary")
    bestRow.CompareMode = vbTextCompare

    Dim bestScore As Object: Set bestScore = CreateObject("Scripting.Dictionary")
    bestScore.CompareMode = vbTextCompare

    For r = 1 To UBound(arr, 1)

        If IsEmpty(arr(r, idxHID)) Or IsEmpty(arr(r, idxHDA)) Then GoTo Prox

        Dim dA As Date: dA = Hist_ToDateOnly(arr(r, idxHDA))
        If dA < baseMin Then GoTo Prox

        Dim idH As String: idH = Trim$(CStr(arr(r, idxHID)))
        If Len(idH) = 0 Then GoTo Prox

        Dim score As String
        score = Format$(dA, "yyyymmdd") & "|"

        If idxHDT > 0 And Not IsEmpty(arr(r, idxHDT)) And IsDate(arr(r, idxHDT)) Then
            score = score & Format$(CDate(arr(r, idxHDT)), "yyyymmddhhnnss")
        Else
            score = score & "00000000000000"
        End If

        If Not bestRow.Exists(idH) Then
            bestRow(idH) = r
            bestScore(idH) = score
        Else
            If score > CStr(bestScore(idH)) Then
                bestRow(idH) = r
                bestScore(idH) = score
            End If
        End If

Prox:
    Next r

    Dim k As Variant
    For Each k In bestRow.keys

        Dim idKey As String: idKey = CStr(k)
        If dictMan.Exists(idKey) Then GoTo NextK

        Dim rowH As Long: rowH = CLng(bestRow(idKey))

        Dim lr As ListRow
        Set lr = loMan.ListRows.Add

        lr.Range.Cells(1, idxMID).Value = idKey

        If idxMParent > 0 And idxHParent > 0 Then lr.Range.Cells(1, idxMParent).Value = arr(rowH, idxHParent)
        If idxMCreated > 0 And idxHCreated > 0 Then lr.Range.Cells(1, idxMCreated).Value = arr(rowH, idxHCreated)
        If idxMMat > 0 And idxHMat > 0 Then lr.Range.Cells(1, idxMMat).Value = arr(rowH, idxHMat)
        If idxMNome > 0 And idxHNome > 0 Then lr.Range.Cells(1, idxMNome).Value = arr(rowH, idxHNome)

        dictMan(idKey) = True

NextK:
    Next k

    Exit Sub

Trata:
    MsgBox "Erro em Manual_IncluirIDs_PeloHistorico_0202: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ===========================================================
' HELPERS
' ===========================================================

Private Function GetTableByName(ByVal tblName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tblName, vbTextCompare) = 0 Then
                Set GetTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set GetTableByName = Nothing
End Function

Private Function GetListColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    GetListColIndex = 0
    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(colName), vbTextCompare) = 0 Then
            GetListColIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Sub EnsureListColumn(ByVal lo As ListObject, ByVal colName As String)
    On Error Resume Next
    Dim tmp As ListColumn
    Set tmp = lo.ListColumns(colName)
    On Error GoTo 0
    If tmp Is Nothing Then lo.ListColumns.Add.Name = colName
End Sub

Private Sub ApplyTextFormatColumn(ByVal lo As ListObject, ByVal colName As String)
    On Error Resume Next
    lo.ListColumns(colName).Range.NumberFormat = "@"
    On Error GoTo 0
End Sub

Private Function IsEmptyOrBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then IsEmptyOrBlank = True: Exit Function
    If Len(Trim$(CStr(v))) = 0 Then IsEmptyOrBlank = True: Exit Function
    IsEmptyOrBlank = False
End Function

Private Function GetLatestHistoricoDateOrToday() As Date
    On Error GoTo Fallback
    Dim loHist As ListObject
    Set loHist = GetTableByName(TB_HIST)
    If loHist Is Nothing Then GoTo Fallback
    If loHist.DataBodyRange Is Nothing Then GoTo Fallback
    Dim idxDA As Long: idxDA = GetListColIndex(loHist, COL_DATAARQUIVO)
    If idxDA = 0 Then GoTo Fallback
    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long, best As Date: best = 0
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDA)) Then
            Dim d As Date: d = DateValue(CDate(arr(r, idxDA)))
            If d > best Then best = d
        End If
    Next r
    If best = 0 Then GoTo Fallback
    GetLatestHistoricoDateOrToday = best
    Exit Function
Fallback:
    GetLatestHistoricoDateOrToday = Date
End Function

Public Sub Sincronizar_RegistrosContatos_Manual()
    On Error GoTo Trata
    Dim wsMan As Worksheet, loMan As ListObject
    Set wsMan = ThisWorkbook.Worksheets(SH_MANUAL)
    Set loMan = wsMan.ListObjects(TBL_MANUAL)
    If loMan Is Nothing Then Exit Sub
    If loMan.DataBodyRange Is Nothing Then Exit Sub
    ApplyTextFormatColumn loMan, COL_ID
    Exit Sub
Trata:
    MsgBox "Erro em Sincronizar_RegistrosContatos_Manual: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Function BuildExtractionDaysDict(ByVal arr As Variant, ByVal idxDA As Long, ByVal dMin As Date, ByVal dMax As Date) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If idxDA = 0 Then Set BuildExtractionDaysDict = dict: Exit Function

    Dim r As Long, d As Date, key As String
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDA)) Then
            d = Hist_ToDateOnly(arr(r, idxDA))
            If d >= dMin And d <= dMax Then
                key = Format(d, "yyyy-mm-dd")
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        End If
    Next r

    Set BuildExtractionDaysDict = dict
End Function

Private Function IsExtractionDay(ByVal dictDays As Object, ByVal d As Variant) As Boolean
    On Error GoTo Fim
    If dictDays Is Nothing Then GoTo Fim
    If Not IsDate(d) Then GoTo Fim
    IsExtractionDay = dictDays.Exists(Format(DateValue(CDate(d)), "yyyy-mm-dd"))
    Exit Function
Fim:
    IsExtractionDay = False
End Function

Private Function GetExtractionDatesSorted(ByVal arr As Variant, ByVal idxDA As Long, ByVal dMin As Date, ByVal dMax As Date) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If idxDA = 0 Then Exit Function

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDA)) Then
            Dim d As Date: d = Hist_ToDateOnly(arr(r, idxDA))
            If d >= dMin And d <= dMax Then
                dict(Format(d, "yyyy-mm-dd")) = d
            End If
        End If
    Next r

    If dict.Count = 0 Then Exit Function

    Dim keys As Variant: keys = dict.keys
    Dim out() As Date
    ReDim out(0 To UBound(keys))

    Dim i As Long
    For i = 0 To UBound(keys)
        out(i) = CDate(dict(keys(i)))
    Next i

    Dim j As Long, tmp As Date
    For i = LBound(out) To UBound(out) - 1
        For j = i + 1 To UBound(out)
            If out(j) < out(i) Then
                tmp = out(i)
                out(i) = out(j)
                out(j) = tmp
            End If
        Next j
    Next i

    GetExtractionDatesSorted = out
End Function

Private Function Hist_ToDateOnly(ByVal v As Variant) As Date
    On Error GoTo Fim
    If IsDate(v) Then
        Hist_ToDateOnly = DateValue(CDate(v))
        Exit Function
    End If
Fim:
    Hist_ToDateOnly = 0
End Function

Private Function GetListColIndexFlex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    Dim a As String, b As String
    a = NormalizeColName(colName)

    For Each lc In lo.ListColumns
        b = NormalizeColName(lc.Name)
        If a = b Then
            GetListColIndexFlex = lc.Index
            Exit Function
        End If
    Next lc
    GetListColIndexFlex = 0
End Function

Private Function NormalizeColName(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "_", "")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    NormalizeColName = s
End Function
