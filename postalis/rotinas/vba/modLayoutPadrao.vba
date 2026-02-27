Option Explicit

' =====================================================================================
' MÓDULO: modLayoutPadrao
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   Aplicar um layout visual padronizado nas abas principais do sistema, ajustando alturas
'   de linha e executando AutoFit de forma inteligente (amostragem), evitando o AutoFit
'   pesado na planilha inteira.
'
' ORIGEM DOS DADOS:
'   - Abas do sistema:
'       * Extração Bitrix
'       * Registros Contatos
'       * Registros Contatos Manual
'   - Tabelas estruturadas (ListObjects) presentes nessas abas
'
' DESTINO:
'   - Formatação visual das abas (RowHeight) e largura de colunas (AutoFit) nas tabelas
'
' COMPORTAMENTO TÉCNICO:
'   - Define alturas padrão:
'       * Linha 1: ALTURA_LINHA_1
'       * Linha 2: ALTURA_LINHA_2
'       * Linhas 3 até última linha: ALTURA_PADRAO
'   - AutoFit “leve”:
'       * Ajusta a largura das colunas usando somente o cabeçalho + primeiras N linhas
'         (AUTOFIT_SAMPLE_ROWS) de cada tabela (ListObject)
'   - Controle diário:
'       * Salva a data do último AutoFit em um Name (NM_LAST_AUTOFIT)
'       * Em execução normal, roda AutoFit somente 1x por dia (salvo quando forçado)
'
' MACROS PÚBLICAS:
'   - AplicarLayoutPadraoTabelas: aplica layout e AutoFit apenas se ainda não rodou hoje
'   - AplicarLayoutPadraoTabelas_Forcar: aplica layout e força AutoFit mesmo que já tenha rodado hoje
'
' DEPENDÊNCIAS:
'   - Chamado por: modAplicarFormulasPendencias.Rotina_Sincronizar_E_Recalcular
'   - Usa Names do Workbook:
'       * RD_LastAutoFitDate (criado automaticamente se não existir)
'
' VERSÃO: 23/02/2026
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private Const ALTURA_LINHA_1 As Double = 40
Private Const ALTURA_LINHA_2 As Double = 30
Private Const ALTURA_PADRAO  As Double = 24

Private Const AUTOFIT_SAMPLE_ROWS As Long = 50
Private Const NM_LAST_AUTOFIT As String = "RD_LastAutoFitDate"

Public Sub AplicarLayoutPadraoTabelas()
    AplicarLayoutPadraoTabelas_Int False
End Sub

Public Sub AplicarLayoutPadraoTabelas_Forcar()
    AplicarLayoutPadraoTabelas_Int True
End Sub

Private Sub AplicarLayoutPadraoTabelas_Int(ByVal forceAutoFit As Boolean)

    On Error GoTo Trata

    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim doAutoFit As Boolean
    doAutoFit = forceAutoFit Or Not JaRodouAutoFitHoje()

    Dim nome As Variant
    For Each nome In Array("Extração Bitrix", "Registros Contatos", "Registros Contatos Manual")
        ApplyLayoutToSheet CStr(nome), doAutoFit
    Next nome

    If doAutoFit Then SalvarAutoFitHoje

Fim:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Trata:
    Resume Fim
End Sub

Private Sub ApplyLayoutToSheet(ByVal sheetName As String, ByVal doAutoFit As Boolean)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    On Error Resume Next
    ws.Rows(1).RowHeight = ALTURA_LINHA_1
    If lastRow >= 2 Then ws.Rows(2).RowHeight = ALTURA_LINHA_2
    If lastRow >= 3 Then ws.Rows("3:" & lastRow).RowHeight = ALTURA_PADRAO
    On Error GoTo 0

    If Not doAutoFit Then Exit Sub

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        AutoFitListObject_Sample lo, AUTOFIT_SAMPLE_ROWS
    Next lo

End Sub

Private Sub AutoFitListObject_Sample(ByVal lo As ListObject, ByVal sampleRows As Long)
    On Error GoTo Fim
    If lo Is Nothing Then Exit Sub

    Dim rngSample As Range
    Set rngSample = GetListObjectSampleRange(lo, sampleRows)
    If rngSample Is Nothing Then GoTo Fim

    rngSample.EntireColumn.AutoFit

Fim:
End Sub

Private Function GetListObjectSampleRange(ByVal lo As ListObject, ByVal sampleRows As Long) As Range

    On Error GoTo Fim

    Dim rngH As Range, rngB As Range
    Set rngH = lo.HeaderRowRange

    If lo.DataBodyRange Is Nothing Then
        Set GetListObjectSampleRange = rngH
        Exit Function
    End If

    Dim n As Long
    n = lo.DataBodyRange.Rows.Count
    If n <= 0 Then
        Set GetListObjectSampleRange = rngH
        Exit Function
    End If

    If sampleRows <= 0 Then sampleRows = 1
    If n > sampleRows Then n = sampleRows

    Set rngB = lo.DataBodyRange.Rows(1).Resize(n)

    Set GetListObjectSampleRange = Union(rngH, rngB)
    Exit Function

Fim:
    Set GetListObjectSampleRange = Nothing
End Function

Private Function JaRodouAutoFitHoje() As Boolean
    On Error GoTo Fim

    Dim v As Variant
    v = GetNameValue(NM_LAST_AUTOFIT)

    If IsDate(v) Then
        JaRodouAutoFitHoje = (DateValue(CDate(v)) = Date)
        Exit Function
    End If

Fim:
    JaRodouAutoFitHoje = False
End Function

Private Sub SalvarAutoFitHoje()
    On Error Resume Next
    SetNameValue NM_LAST_AUTOFIT, CDate(Date)
    On Error GoTo 0
End Sub

Private Function GetNameValue(ByVal nm As String) As Variant
    On Error GoTo Fim
    GetNameValue = ThisWorkbook.Names(nm).RefersToRange.Value
    Exit Function
Fim:
    GetNameValue = Empty
End Function

Private Sub SetNameValue(ByVal nm As String, ByVal v As Variant)
    On Error GoTo Cria
    ThisWorkbook.Names(nm).RefersToRange.Value = v
    Exit Sub

Cria:
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    ThisWorkbook.Names.Add Name:=nm, RefersTo:=ws.Range("A1")
    ThisWorkbook.Names(nm).RefersToRange.Value = v
    On Error GoTo 0
End Sub
