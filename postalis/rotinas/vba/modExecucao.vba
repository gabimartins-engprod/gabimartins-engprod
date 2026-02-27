Option Explicit

' =====================================================================================
' MÓDULO: modExecucao
' SISTEMA: Rotina_Dados – Postalis
' BOTÃO ÚNICO: AtualizarTudo_E_Sincronizar
'
' FINALIDADE:
'   Orquestrar a execução completa da rotina diária (Power Query + sincronizações + log),
'   garantindo que as etapas rodem na ordem correta, sem duplicidade por clique e com
'   robustez para dias sem extração.
'
' ORIGEM DOS DADOS:
'   - Power Query (RefreshAll) que atualiza as tabelas/consultas do arquivo
'   - Tabela do Histórico Bitrix: tblHistoricoBitrix (aba "Histórico Bitrix")
'
' DESTINO:
'   - Tabela de Log: tblLogIDs (aba "Log_Sincronizacao")
'   - Tabelas finais/suporte atualizadas pelas rotinas chamadas:
'       * Correção_Automática (fórmulas espelhadas)
'       * Registros_Contatos_Final (Dentro/Fora + datas)
'       * Registros_Contatos_Manual (inclusão de IDs + backfill leve de datas)
'   - Atualização de tabelas dinâmicas (pivôs) ao final
'
' PATCH 19/02/2026 (ROBUSTEZ - DEFINITIVO):
'   - DataArquivo (date) vira a referência oficial do "dia" (calendário).
'   - DataHoraExtracao (datetime) fica APENAS para escolher a última extração daquele dia.
'   - dictDays / maxHistDate usam DataArquivo (quando existir).
'   - LoadIdsForLastExtractionOfDay_Robusto filtra por DataArquivo e:
'       * Se existir DataHoraExtracao no dia: pega SOMENTE a maior DataHoraExtracao do dia.
'       * Se NÃO existir DataHoraExtracao no dia: pega TODAS as linhas do dia (por DataArquivo).
'   - ResolveBaseline volta a carregar prevSet (evita “todo mundo entrou” após lacunas).
'
' PATCH 25/02/2026 (FIX PRIMEIRO DIA COM EXTRAÇÃO):
'   - Se startD tiver dias sem extração no início, o 1º dia QUE TIVER extração não pode listar
'     todos os IDs como ENTROU (prevSet vazio).
'   - Agora, se lastAvailableDay=0 e hasExtraction=True:
'       * Trata como "primeiro dia com extração" => só INICIO/RESUMO/FIM e inicializa prevSet.
'
' PATCH 25/02/2026 (OPÇÃO A + CORREÇÕES):
'   - Botão com “run once por etapa” (remove duplicidades por clique).
'   - Garante sincronização/backfill (corrige Data Entrada faltando).
'   - Preserva/reaplica TotalDia no Log_Sincronizacao (corrige coluna em branco).
'   - Atualiza pivôs no final com recálculo normalizado.
'
' FLUXO DA ROTINA (ordem do botão):
'   1) (Opcional) Força Power Query a não rodar em background: modPQPowerQuery.PQ_BackgroundOff
'   2) RefreshAll + espera consultas assíncronas terminarem
'   3) Aplicar fórmulas na Correção_Automática + layout: modAplicarFormulasPendencias.Rotina_Sincronizar_E_Recalcular
'   4) Sincronizar Registros_Contatos_Final (Dentro/Fora + datas): modSincronizarContatos.Sincronizar_AteX_RegistrosContatos
'   5) Atualizar Registros_Contatos_Manual (incluir IDs + histórico): Manual_IncluirNovosIDs_PeloFinal + Manual_IncluirIDs_PeloHistorico_0202
'   6) Backfill leve (pendentes): Manual_Backfill_Rapido_0202
'   7) Gerar Log diário (inclui dias sem extração): BuildLog_FromHistorico_FillMissingDays (a partir de 01/02/2026)
'   8) Recalcular TotalDia (somente linhas RESUMO) e atualizar pivôs
'
' DEPENDÊNCIAS (chamadas):
'   - modPQPowerQuery.PQ_BackgroundOff (opcional)
'   - modAplicarFormulasPendencias.Rotina_Sincronizar_E_Recalcular
'   - modSincronizarContatos.Sincronizar_AteX_RegistrosContatos
'   - modSincronizarContatos.Manual_IncluirNovosIDs_PeloFinal
'   - modSincronizarContatos.Manual_IncluirIDs_PeloHistorico_0202
'   - modSincronizarContatos.Manual_Backfill_Rapido_0202
'
' TABELAS/ABAS REFERENCIADAS:
'   - Log_Sincronizacao / tblLogIDs
'   - Histórico Bitrix / tblHistoricoBitrix
'
' VERSÃO: 25/02/2026 (Opção A + Fixes)
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private isRunning As Boolean
Private gRanSteps As Object ' Scripting.Dictionary (late binding)

Private Const SH_LOG As String = "Log_Sincronizacao"
Private Const TBL_LOG As String = "tblLogIDs"

Private Const SH_HIST As String = "Histórico Bitrix"
Private Const TB_HIST As String = "tblHistoricoBitrix"

Private Const COL_ID As String = "ID"
Private Const COL_DATAARQUIVO As String = "DataArquivo"               ' Date (base do dia)
Private Const COL_DATAHORA_EXTRACAO As String = "DataHoraExtracao"    ' DateTime (última extração do dia)

Private Const LOG_COL_TOTALDIA As String = "TotalDia"                 ' Nome EXATO da coluna no tblLogIDs

' (Opcional) se existir no seu arquivo: força PQ a não rodar em background
Private Const MACRO_PQ_BACKGROUND_OFF As String = "modPQPowerQuery.PQ_BackgroundOff"

Public Sub AtualizarTudo_E_Sincronizar()
    If isRunning Then Exit Sub
    isRunning = True
    On Error GoTo TrataErro

    Dim prevEvents As Boolean, prevScreen As Boolean, prevDisplayStatus As Boolean
    Dim prevStatusBar As Variant, prevCalc As XlCalculation

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevStatusBar = Application.StatusBar
    prevDisplayStatus = Application.DisplayStatusBar
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual

    InitRunCache

    Application.StatusBar = "Preparando Power Query..."
    RunStepOnce "PQ_BackgroundOff", "RUN", MACRO_PQ_BACKGROUND_OFF, True

    Application.StatusBar = "Atualizando Power Query..."
    RunStepOnce "RefreshAllAndWait", "REFRESH", "", False

    Application.StatusBar = "Sincronização interna: fórmulas/layout..."
    DoEvents
    RunStepOnce "Rotina_Sincronizar_E_Recalcular", "CALL", "modAplicarFormulasPendencias.Rotina_Sincronizar_E_Recalcular", False

    Application.StatusBar = "Final: sincronizando Dentro/Fora + datas..."
    DoEvents
    RunStepOnce "Sincronizar_AteX_RegistrosContatos", "CALL", "modSincronizarContatos.Sincronizar_AteX_RegistrosContatos", False

    Application.StatusBar = "Manual: incluir novos IDs..."
    DoEvents
    RunStepOnce "Manual_IncluirNovosIDs_PeloFinal", "CALL", "modSincronizarContatos.Manual_IncluirNovosIDs_PeloFinal", False
    RunStepOnce "Manual_IncluirIDs_PeloHistorico_0202", "CALL", "modSincronizarContatos.Manual_IncluirIDs_PeloHistorico_0202", False

    Application.StatusBar = "Manual: backfill de datas..."
    DoEvents
    RunStepOnce "Manual_Backfill_Rapido_0202", "CALL", "modSincronizarContatos.Manual_Backfill_Rapido_0202", False

    Application.StatusBar = "Gerando Log Histórico (dias faltantes)..."
    DoEvents
    RunStepOnce "BuildLog_FromHistorico", "LOG", "", False

    Application.StatusBar = "Recalculando (TotalDia) + Ajustes finais..."
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    FixTotalDia_IfExists

    Application.StatusBar = "Atualizando Tabelas Dinâmicas..."
    RefreshAllPivots

Saida:
    Application.StatusBar = prevStatusBar
    Application.DisplayStatusBar = prevDisplayStatus
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Set gRanSteps = Nothing
    isRunning = False
    Exit Sub

TrataErro:
    Application.StatusBar = "Erro: " & Err.Description
    Resume Saida
End Sub

' =====================================================================================
' RUN-ONCE POR ETAPA (anti-duplicidade por clique)
' =====================================================================================
Private Sub InitRunCache()
    Set gRanSteps = CreateObject("Scripting.Dictionary")
    gRanSteps.CompareMode = 1 ' vbTextCompare
End Sub

Private Sub RunStepOnce(ByVal stepKey As String, ByVal mode As String, ByVal macroFullName As String, ByVal isOptional As Boolean)
    If gRanSteps Is Nothing Then InitRunCache
    If gRanSteps.Exists(stepKey) Then Exit Sub
    gRanSteps.Add stepKey, True
    Select Case UCase$(mode)
        Case "RUN"
            If isOptional Then SafeRunOptional macroFullName Else SafeRunRequired macroFullName
        Case "CALL"
            If isOptional Then SafeRunOptional macroFullName Else SafeRunRequired macroFullName
        Case "REFRESH"
            RefreshAllAndWait
        Case "LOG"
            BuildLog_FromHistorico_FillMissingDays DateSerial(2026, 2, 1)
        Case Else
            ' nada
    End Select
End Sub

Private Sub RefreshAllAndWait()
    ThisWorkbook.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
    DoEvents
End Sub

Private Sub SafeRunOptional(ByVal macroFullName As String)
    On Error Resume Next
    If Len(Trim$(macroFullName)) = 0 Then Exit Sub
    Application.Run macroFullName
    On Error GoTo 0
End Sub

Private Sub SafeRunRequired(ByVal macroFullName As String)
    On Error GoTo Trata
    If Len(Trim$(macroFullName)) = 0 Then Exit Sub
    Application.Run macroFullName
    Exit Sub
Trata:
    Err.Raise vbObjectError + 513, "modExecucao", "Não foi possível executar: [" & macroFullName & "] -> " & Err.Description
End Sub

' =====================================================================================
' FIX: TOTALDIA (VBA) — preenche SOMENTE nas linhas RESUMO (como era antes)
' =====================================================================================
Private Sub FixTotalDia_IfExists()
    On Error GoTo Fim

    Dim loLog As ListObject
    Set loLog = FindTableByName(ThisWorkbook, TBL_LOG)
    If loLog Is Nothing Then GoTo Fim
    If loLog.DataBodyRange Is Nothing Then GoTo Fim

    ' Garante coluna TotalDia
    Dim lcTotal As ListColumn
    On Error Resume Next
    Set lcTotal = loLog.ListColumns(LOG_COL_TOTALDIA)
    On Error GoTo Fim
    If lcTotal Is Nothing Then
        Set lcTotal = loLog.ListColumns.Add
        lcTotal.Name = LOG_COL_TOTALDIA
    End If

    Dim idxAcao As Long, idxObs As Long, idxTotal As Long
    idxAcao = GetListColIndex(loLog, "Acao")
    idxObs = GetListColIndex(loLog, "Obs")
    idxTotal = lcTotal.Index
    If idxAcao = 0 Or idxObs = 0 Or idxTotal = 0 Then GoTo Fim

    Dim arr As Variant
    arr = loLog.DataBodyRange.Value

    Dim r As Long, acao As String, obs As String, tot As Long
    For r = 1 To UBound(arr, 1)
        acao = UCase$(Trim$(CStr(arr(r, idxAcao))))
        If acao = "RESUMO" Then
            obs = CStr(arr(r, idxObs))
            tot = ExtractTotalNoDia(obs)
            If tot >= 0 Then
                arr(r, idxTotal) = tot
            Else
                arr(r, idxTotal) = vbNullString
            End If
        Else
            arr(r, idxTotal) = vbNullString
        End If
    Next r

    loLog.DataBodyRange.Value = arr
    lcTotal.DataBodyRange.NumberFormat = "0"

Fim:
End Sub

' =====================================================================================
' Extrai N de "... Total no dia: N" (retorna -1 se não achar)
' =====================================================================================
Private Function ExtractTotalNoDia(ByVal s As String) As Long
    On Error GoTo Falha

    Dim p As Long, t As String, i As Long
    ExtractTotalNoDia = -1

    p = InStr(1, s, "Total no dia:", vbTextCompare)
    If p = 0 Then Exit Function

    t = Trim$(Mid$(s, p + Len("Total no dia:")))
    If Len(t) = 0 Then Exit Function

    ' pega só os dígitos iniciais
    For i = 1 To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then
            t = Left$(t, i - 1)
            Exit For
        End If
    Next i

    If Len(t) = 0 Then Exit Function
    ExtractTotalNoDia = CLng(t)
    Exit Function

Falha:
    ExtractTotalNoDia = -1
End Function

' =====================================================================================
' LOG HISTÓRICO (COMPLETA DIAS FALTANTES) - DIA POR DataArquivo + ÚLTIMA EXTRAÇÃO POR DataHoraExtracao
' =====================================================================================

Private Sub BuildLog_FromHistorico_FillMissingDays(ByVal startD As Date)

    Dim loLog As ListObject: Set loLog = EnsureLogTable()
    Dim loHist As ListObject: Set loHist = FindTableByName(ThisWorkbook, TB_HIST)

    If loLog Is Nothing Then Exit Sub
    If loHist Is Nothing Then Exit Sub
    If loHist.DataBodyRange Is Nothing Then Exit Sub

    ApplyTextFormatColumn loLog, 3
    ApplyTextFormatListObject loHist, COL_ID

    Dim idxID As Long: idxID = GetListColIndex(loHist, COL_ID)
    Dim idxDA As Long: idxDA = GetListColIndex(loHist, COL_DATAARQUIVO)
    Dim idxDT As Long: idxDT = GetListColIndex(loHist, COL_DATAHORA_EXTRACAO)

    If idxID = 0 Then Exit Sub
    If idxDA = 0 Then Exit Sub

    Dim firstSeen As Object: Set firstSeen = CreateObject("Scripting.Dictionary")
    firstSeen.CompareMode = vbTextCompare

    Dim arrFS As Variant, rFS As Long
    arrFS = loHist.DataBodyRange.Value

    For rFS = 1 To UBound(arrFS, 1)
        If IsEmpty(arrFS(rFS, idxID)) Then GoTo NextFS
        If IsEmpty(arrFS(rFS, idxDA)) Then GoTo NextFS
        Dim idFS As String: idFS = Trim$(CStr(arrFS(rFS, idxID)))
        If Len(idFS) = 0 Then GoTo NextFS
        On Error Resume Next
        Dim dFS As Date: dFS = DateValue(CDate(arrFS(rFS, idxDA)))
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFS
        Err.Clear: On Error GoTo 0
        If Not firstSeen.Exists(idFS) Then
            firstSeen(idFS) = dFS
        Else
            If dFS < CDate(firstSeen(idFS)) Then firstSeen(idFS) = dFS
        End If
NextFS:
    Next rFS

    Dim maxHistDate As Date
    maxHistDate = GetMaxDateInHistorico_Robusto(loHist, idxDA, idxDT)
    If maxHistDate = 0 Then Exit Sub
    If maxHistDate < startD Then Exit Sub

    Dim lastProcessed As Date
    lastProcessed = GetLastProcessedDate(loLog)

    Dim firstToProcess As Date
    If lastProcessed = 0 Then
        firstToProcess = startD
    Else
        firstToProcess = DateAdd("d", 1, lastProcessed)
    End If
    If firstToProcess > maxHistDate Then Exit Sub

    Dim dictLoggedDays As Object
    Set dictLoggedDays = BuildProcessedDaysDict(loLog)

    Dim dictDays As Object
    Set dictDays = BuildDaysWithExtractionDict_Robusto(loHist, startD, maxHistDate, idxDA)

    Dim prevSet As Object: Set prevSet = CreateObject("Scripting.Dictionary"): prevSet.CompareMode = vbTextCompare
    Dim everSeen As Object: Set everSeen = CreateObject("Scripting.Dictionary"): everSeen.CompareMode = vbTextCompare
    Dim lastAvailableDay As Date: lastAvailableDay = 0
    ResolveBaseline loHist, dictDays, startD, firstToProcess, prevSet, everSeen, lastAvailableDay, idxID, idxDA, idxDT

    Dim d As Date
    For d = firstToProcess To maxHistDate

        Dim hasExtraction As Boolean
        hasExtraction = dictDays.Exists(Format(d, "yyyy-mm-dd"))

        If DayAlreadyLogged(dictLoggedDays, d) Then
            If hasExtraction Then
                Dim tmpSet As Object: Set tmpSet = CreateObject("Scripting.Dictionary"): tmpSet.CompareMode = vbTextCompare
                Dim tmpStamp As Date
                LoadIdsForLastExtractionOfDay_Robusto loHist, d, tmpSet, tmpStamp, idxID, idxDA, idxDT
                If tmpStamp <> 0 Then
                    prevSet.RemoveAll
                    CopyDict tmpSet, prevSet
                    MarkEverSeen tmpSet, everSeen
                    lastAvailableDay = d
                End If
            End If
            GoTo ProximoDia
        End If

        Dim prevLabel As String
        If lastAvailableDay = 0 Then prevLabel = "(primeiro dia com extração)" Else prevLabel = Format(lastAvailableDay, "dd/mm/yyyy")

        If Not hasExtraction Then
            Dim stampNo As Date
            stampNo = DateSerial(Year(d), Month(d), Day(d))
            AppendLog loLog, stampNo, "INICIO", "", "Comparativo Histórico Bitrix: " & prevLabel & " \ " & Format(d, "dd/mm/yyyy")
            AppendLog loLog, stampNo, "RESUMO", "", "DataArquivo: " & Format(d, "dd/mm/yyyy") & "; Não houve extração automática nessa data. Verificar o motivo com o Six."
            AppendLog loLog, stampNo, "FIM", "", "Comparativo concluído."
        Else
            Dim curSet As Object: Set curSet = CreateObject("Scripting.Dictionary"): curSet.CompareMode = vbTextCompare
            Dim stampDT As Date
            LoadIdsForLastExtractionOfDay_Robusto loHist, d, curSet, stampDT, idxID, idxDA, idxDT
            If stampDT = 0 Then GoTo ProximoDia

            If lastAvailableDay = 0 Then
                AppendLog loLog, stampDT, "INICIO", "", "Comparativo Histórico Bitrix: (primeiro dia com extração) \ " & Format(d, "dd/mm/yyyy")
                AppendLog loLog, stampDT, "RESUMO", "", "DataArquivo: " & Format(d, "dd/mm/yyyy") & "; Total no dia: " & curSet.Count
                AppendLog loLog, stampDT, "FIM", "", "Comparativo concluído."
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
                    If firstSeen.Exists(CStr(k)) Then
                        If CDate(firstSeen(CStr(k))) < d Then retornou.Add CStr(k) Else entrou.Add CStr(k)
                    Else
                        entrou.Add CStr(k)
                    End If
                    everSeen(CStr(k)) = True
                End If
            Next k

            For Each k In prevSet.Keys
                If Not curSet.Exists(CStr(k)) Then saiu.Add CStr(k)
            Next k

            AppendLog loLog, stampDT, "INICIO", "", "Comparativo Histórico Bitrix: " & prevLabel & " \ " & Format(d, "dd/mm/yyyy")
            If retornou.Count > 0 Then WriteIds loLog, stampDT, "RETORNOU", retornou, "Voltou na extração (" & Format(d, "dd/mm/yyyy") & ")"
            If saiu.Count > 0 Then WriteIds loLog, stampDT, "SAIU", saiu, "Saiu da extração (" & Format(d, "dd/mm/yyyy") & ")"
            If entrou.Count > 0 Then WriteIds loLog, stampDT, "ENTROU", entrou, "Entrou na extração (" & Format(d, "dd/mm/yyyy") & ")"

            AppendLog loLog, stampDT, "RESUMO", "", "DataArquivo: " & Format(d, "dd/mm/yyyy") & "; ENTROU: " & entrou.Count & "; RETORNOU: " & retornou.Count & "; SAIU: " & saiu.Count & "; Total no dia: " & curSet.Count
            AppendLog loLog, stampDT, "FIM", "", "Comparativo concluído."

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

Private Sub ResolveBaseline(ByVal loHist As ListObject, ByVal dictDays As Object, ByVal startD As Date, ByVal firstToProcess As Date, ByRef prevSet As Object, ByRef everSeen As Object, ByRef lastAvailableDay As Date, ByVal idxID As Long, ByVal idxDA As Long, ByVal idxDT As Long)
    prevSet.RemoveAll
    everSeen.RemoveAll
    lastAvailableDay = 0
    If firstToProcess <= startD Then Exit Sub
    Dim scan As Date: scan = DateAdd("d", -1, firstToProcess)
    Do While scan >= startD
        If dictDays.Exists(Format(scan, "yyyy-mm-dd")) Then
            Dim tmpStamp As Date
            LoadIdsForLastExtractionOfDay_Robusto loHist, scan, prevSet, tmpStamp, idxID, idxDA, idxDT
            If tmpStamp <> 0 Then
                MarkEverSeen prevSet, everSeen
                lastAvailableDay = scan
            End If
            Exit Sub
        End If
        scan = DateAdd("d", -1, scan)
    Loop
End Sub

Private Function GetMaxDateInHistorico_Robusto(ByVal loHist As ListObject, ByVal idxDA As Long, ByVal idxDT As Long) As Date
    If loHist.DataBodyRange Is Nothing Then Exit Function
    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long, best As Date: best = 0
    For r = 1 To UBound(arr, 1)
        If idxDA > 0 And Not IsEmpty(arr(r, idxDA)) Then
            On Error Resume Next
            Dim dA As Date: dA = DateValue(CDate(arr(r, idxDA)))
            If Err.Number = 0 Then If dA > best Then best = dA
            Err.Clear: On Error GoTo 0
        End If
    Next r
    If best = 0 And idxDT > 0 Then
        For r = 1 To UBound(arr, 1)
            If Not IsEmpty(arr(r, idxDT)) Then
                On Error Resume Next
                Dim dT As Date: dT = DateValue(CDate(arr(r, idxDT)))
                If Err.Number = 0 Then If dT > best Then best = dT
                Err.Clear: On Error GoTo 0
            End If
        Next r
    End If
    GetMaxDateInHistorico_Robusto = best
End Function

Private Function BuildDaysWithExtractionDict_Robusto(ByVal loHist As ListObject, ByVal d1 As Date, ByVal d2 As Date, ByVal idxDA As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If loHist.DataBodyRange Is Nothing Then Set BuildDaysWithExtractionDict_Robusto = dict: Exit Function
    If idxDA = 0 Then Set BuildDaysWithExtractionDict_Robusto = dict: Exit Function
    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, idxDA)) Then
            On Error Resume Next
            Dim dA As Date: dA = DateValue(CDate(arr(r, idxDA)))
            If Err.Number = 0 Then
                If dA >= d1 And dA <= d2 Then
                    Dim key As String: key = Format(dA, "yyyy-mm-dd")
                    If Not dict.Exists(key) Then dict.Add key, True
                End If
            End If
            Err.Clear: On Error GoTo 0
        End If
    Next r
    Set BuildDaysWithExtractionDict_Robusto = dict
End Function

Private Sub LoadIdsForLastExtractionOfDay_Robusto(ByVal loHist As ListObject, ByVal dayD As Date, ByRef dictOut As Object, ByRef stampOut As Date, ByVal idxID As Long, ByVal idxDA As Long, ByVal idxDT As Long)
    dictOut.RemoveAll
    stampOut = 0
    If loHist.DataBodyRange Is Nothing Then Exit Sub
    If idxID = 0 Or idxDA = 0 Then Exit Sub
    Dim arr As Variant: arr = loHist.DataBodyRange.Value
    Dim r As Long
    Dim maxDT As Date: maxDT = 0
    Dim foundDT As Boolean: foundDT = False
    If idxDT > 0 Then
        For r = 1 To UBound(arr, 1)
            If Not IsEmpty(arr(r, idxDA)) Then
                On Error Resume Next
                Dim dA As Date: dA = DateValue(CDate(arr(r, idxDA)))
                If Err.Number = 0 And dA = dayD Then
                    If Not IsEmpty(arr(r, idxDT)) Then
                        Dim dtV As Date: dtV = CDate(arr(r, idxDT))
                        If Err.Number = 0 Then
                            foundDT = True
                            If dtV > maxDT Then maxDT = dtV
                        End If
                    End If
                End If
                Err.Clear: On Error GoTo 0
            End If
        Next r
    End If
    For r = 1 To UBound(arr, 1)
        If IsEmpty(arr(r, idxDA)) Or IsEmpty(arr(r, idxID)) Then GoTo ProxR
        On Error Resume Next
        Dim dA2 As Date: dA2 = DateValue(CDate(arr(r, idxDA)))
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo ProxR
        Err.Clear: On Error GoTo 0
        If dA2 <> dayD Then GoTo ProxR
        If foundDT Then
            If idxDT = 0 Or IsEmpty(arr(r, idxDT)) Then GoTo ProxR
            On Error Resume Next
            Dim dt2 As Date: dt2 = CDate(arr(r, idxDT))
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo ProxR
            Err.Clear: On Error GoTo 0
            If dt2 <> maxDT Then GoTo ProxR
        End If
        Dim idS As String: idS = Trim$(CStr(arr(r, idxID)))
        If Len(idS) > 0 Then dictOut(idS) = True
ProxR:
    Next r
    If foundDT Then stampOut = maxDT Else stampOut = DateSerial(Year(dayD), Month(dayD), Day(dayD))
End Sub

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
    ApplyTextFormatColumn lo, 3
    Set EnsureLogTable = lo
End Function

Private Sub AppendLog(ByVal loLog As ListObject, ByVal dT As Date, ByVal acao As String, ByVal idVal As String, ByVal obs As String)
    Dim lr As ListRow
    Set lr = loLog.ListRows.Add
    lr.Range.Cells(1, 1).Value = dT
    lr.Range.Cells(1, 2).Value = acao
    lr.Range.Cells(1, 3).Value = CStr(idVal)
    lr.Range.Cells(1, 4).Value = obs
End Sub

Private Sub WriteIds(ByVal loLog As ListObject, ByVal dT As Date, ByVal acao As String, ByVal col As Collection, ByVal obs As String)
    Dim i As Long
    For i = 1 To col.Count
        AppendLog loLog, dT, acao, CStr(col(i)), obs
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
    Dim v As String: v = Trim$(Mid$(s, p + 1))
    If Len(v) >= 10 Then
        ParseMarcoDate = DateSerial(CInt(Mid$(v, 1, 4)), CInt(Mid$(v, 6, 2)), CInt(Mid$(v, 9, 2)))
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
            Dim d As Date: d = ParseDataArquivoFromObs(CStr(arr(r, 4)))
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
    Dim s As String: s = Trim$(Mid$(obs, p + Len("DataArquivo:")))
    If Len(s) < 10 Then GoTo Fim
    Dim dd As Integer, mm As Integer, yy As Integer
    dd = CInt(Mid$(s, 1, 2))
    mm = CInt(Mid$(s, 4, 2))
    yy = CInt(Mid$(s, 7, 4))
    ParseDataArquivoFromObs = DateSerial(yy, mm, dd)
    Exit Function
Fim:
    ParseDataArquivoFromObs = 0
End Function

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

Private Sub ApplyTextFormatColumn(ByVal lo As ListObject, ByVal colIndex As Long)
    On Error Resume Next
    lo.ListColumns(colIndex).Range.NumberFormat = "@"
    On Error GoTo 0
End Sub

Private Sub ApplyTextFormatListObject(ByVal lo As ListObject, ByVal colName As String)
    On Error Resume Next
    lo.ListColumns(colName).Range.NumberFormat = "@"
    On Error GoTo 0
End Sub
