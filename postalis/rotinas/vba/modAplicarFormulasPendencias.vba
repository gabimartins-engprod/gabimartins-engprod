Option Explicit

' =====================================================================================
' MÓDULO: modAplicarFormulasPendencias
' SISTEMA: Rotina_Dados – Postalis
'
' FINALIDADE:
'   Espelhar, na tabela Correção_Automática, os campos manuais de contato preenchidos em
'   Registros_Contatos_Final, através de colunas calculadas (PROCX + LET), mantendo:
'     - ID tratado como TEXTO (evita perda/alteração de IDs numéricos)
'     - Robustez para tabela vazia
'     - Restauração do estado do Excel (cálculo/eventos/tela/statusbar)
'
' ORIGEM DOS DADOS:
'   - Tabela Registros_Contatos_Final (base oficial com campos manuais preenchidos)
'
' DESTINO:
'   - Tabela Correção_Automática (colunas calculadas locais por linha, espelhando dados manuais)
'
' COLUNAS ESPELHADAS (destino -> origem):
'   - E-mail -> E-mail
'   - Telefone -> Telefone
'   - Data_1 -> Data_1
'   - OBS_1 -> OBS_1
'   - Responsável_1 -> Responsável_1
'   - Data_2 -> Data_2
'   - OBS_2 -> OBS_2
'   - Responsável_2 -> Responsável_2
'   - Data_3 -> Data_3
'   - OBS_3 -> OBS_3
'   - Responsável_3 -> Responsável_3
'   - Observação Extra -> Observação Extra
'   - Documentos Recebidos -> Documentos Recebidos
'
' COMPORTAMENTO TÉCNICO:
'   - Para colunas que contém "Data": converte texto em data (DATA.VALOR) e evita 0/""
'   - Para demais colunas: mantém texto/valor, evitando 0/""
'   - Aplica formatação:
'       * Datas: "dd/mmm/aa"
'       * Telefone / E-mail: "@"
'
' OTIMIZAÇÃO (TURBO):
'   - Só reescreve fórmulas quando necessário:
'       * Coluna ainda não existe; ou
'       * Fórmula atual (normalizada) é diferente da fórmula desejada
'
' OPÇÃO A (ORQUESTRAÇÃO):
'   - modExecucao é o responsável por orquestrar o fluxo completo.
'   - Este módulo NÃO chama mais rotinas de sincronização (evita duplicidades).
'   - Rotina_Sincronizar_E_Recalcular mantém apenas:
'       (1) AplicarFormulasPendencias
'       (2) AplicarLayoutPadraoTabelas (modLayoutPadrao)
'
' TABELAS REFERENCIADAS:
'   - Correção_Automática
'   - Registros_Contatos_Final
'
' VERSÃO: 25/02/2026 (Opção A: sem orquestração duplicada)
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private Const TBL_PEND As String = "Correção_Automática"
Private Const TBL_CONT As String = "Registros_Contatos_Final"
Private Const COL_ID As String = "ID"

Private COLS_MAP As Variant

Private Sub InitMap()
    COLS_MAP = Array( _
        Array("E-mail", "E-mail"), _
        Array("Telefone", "Telefone"), _
        Array("Data_1", "Data_1"), _
        Array("OBS_1", "OBS_1"), _
        Array("Responsável_1", "Responsável_1"), _
        Array("Data_2", "Data_2"), _
        Array("OBS_2", "OBS_2"), _
        Array("Responsável_2", "Responsável_2"), _
        Array("Data_3", "Data_3"), _
        Array("OBS_3", "OBS_3"), _
        Array("Responsável_3", "Responsável_3"), _
        Array("Observação Extra", "Observação Extra"), _
        Array("Documentos Recebidos", "Documentos Recebidos") _
    )
End Sub

Public Sub AplicarFormulasPendencias()
    On Error GoTo Tratar

    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim prevStatusBar As Variant

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    InitMap

    Dim loPend As ListObject
    Dim loCont As ListObject
    Dim i As Long, destCol As String, srcCol As String, f As String
    Dim didChange As Boolean

    Set loPend = GetTableByName(TBL_PEND)
    Set loCont = GetTableByName(TBL_CONT)

    If loPend Is Nothing Or loCont Is Nothing Then
        MsgBox "Verifique as tabelas '" & TBL_PEND & "' e '" & TBL_CONT & "'.", vbExclamation
        GoTo Fim
    End If

    If loPend.DataBodyRange Is Nothing Then GoTo Fim

    For i = LBound(COLS_MAP) To UBound(COLS_MAP)

        destCol = CStr(COLS_MAP(i)(0))
        srcCol = CStr(COLS_MAP(i)(1))

        If HasListColumn(loCont, srcCol) Then

            Application.StatusBar = "Fórmulas (Correção_Automática): " & destCol
            DoEvents

            If InStr(1, destCol, "Data", vbTextCompare) > 0 Then
                f = "=LET(" & _
                    "v; PROCX([" & COL_ID & "]; " & loCont.Name & "[" & COL_ID & "]; " & loCont.Name & "[" & srcCol & "]; """"); " & _
                    "SE(OU(v=""""; v=0); """"; SE(ÉTEXTO(v); DATA.VALOR(v); v))" & _
                    ")"

                didChange = EnsureCalculatedColumnLocal_Turbo(loPend, destCol, f)
                If didChange Then ApplyNumberFormatIfData loPend, destCol, "dd/mmm/aa"
            Else
                f = "=LET(" & _
                    "v; PROCX([" & COL_ID & "]; " & loCont.Name & "[" & COL_ID & "]; " & loCont.Name & "[" & srcCol & "]; """"); " & _
                    "SE(OU(v=""""; v=0); """"; v)" & _
                    ")"

                didChange = EnsureCalculatedColumnLocal_Turbo(loPend, destCol, f)

                If didChange Then
                    If LCase$(destCol) = "telefone" Or LCase$(destCol) = "e-mail" Then
                        ApplyNumberFormatIfData loPend, destCol, "@"
                    End If
                End If
            End If

        End If
    Next i

Fim:
    Application.StatusBar = prevStatusBar
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Tratar:
    Application.StatusBar = prevStatusBar
    MsgBox "Erro em AplicarFormulasPendencias: " & Err.Description, vbExclamation
    Resume Fim
End Sub

' Mantida por compatibilidade, mas SEM chamadas de sincronização (Opção A).
Public Sub Rotina_Sincronizar_E_Recalcular()
    On Error GoTo Tratar

    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim prevStatusBar As Variant

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "Correção_Automática: fórmulas..."
    DoEvents
    AplicarFormulasPendencias

    Application.StatusBar = "Layout..."
    DoEvents
    modLayoutPadrao.AplicarLayoutPadraoTabelas

Fim:
    Application.StatusBar = prevStatusBar
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Tratar:
    Application.StatusBar = prevStatusBar
    MsgBox "Erro na rotina: " & Err.Description, vbExclamation
    Resume Fim
End Sub

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
End Function

Private Function HasListColumn(ByVal lo As ListObject, ByVal colName As String) As Boolean
    On Error Resume Next
    HasListColumn = Not lo.ListColumns(colName) Is Nothing
    On Error GoTo 0
End Function

Private Sub ApplyNumberFormatIfData(ByVal lo As ListObject, ByVal colName As String, ByVal fmt As String)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    lo.ListColumns(colName).DataBodyRange.NumberFormat = fmt
    On Error GoTo 0
End Sub

Private Function EnsureCalculatedColumnLocal_Turbo(ByVal lo As ListObject, ByVal colName As String, ByVal formulaLocal As String) As Boolean
    EnsureCalculatedColumnLocal_Turbo = False
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0

    If lc Is Nothing Then
        lo.ListColumns.Add.Name = colName
        Set lc = lo.ListColumns(colName)
        lc.DataBodyRange.FormulaLocal = formulaLocal
        EnsureCalculatedColumnLocal_Turbo = True
        Exit Function
    End If

    Dim cur As String
    On Error Resume Next
    cur = CStr(lc.DataBodyRange.Cells(1, 1).FormulaLocal)
    On Error GoTo 0

    If NormalizeFormula(cur) <> NormalizeFormula(formulaLocal) Then
        lc.DataBodyRange.FormulaLocal = formulaLocal
        EnsureCalculatedColumnLocal_Turbo = True
    End If
End Function

Private Function NormalizeFormula(ByVal s As String) As String
    s = Replace(s, " ", "")
    NormalizeFormula = s
End Function
