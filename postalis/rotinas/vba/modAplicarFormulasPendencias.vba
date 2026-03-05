' =========================
' modAplicarFormulasPendencias
' =========================
Option Explicit

' =====================================================================================
' MÓDULO: modAplicarFormulasPendencias
' SISTEMA: Rotina_Dados – Postalis
' FINALIDADE:
'   - (Arquitetura atual) Correção_Automática NÃO recebe espelhamento de campos do Manual/Final.
'   - Mantido para compatibilidade com modExecucao (Opção A) e para layout padronizado.
'
' OBS:
'   - Este módulo NÃO cria colunas na Correção_Automática.
'   - Se no futuro quiser espelhar algo, o espelhamento deve ser planejado com colunas EXISTENTES.
'
' VERSÃO: 03/03/2026 (blindagem: sem criação de colunas)
' RESPONSÁVEL: Gabi (Postalis)
' =====================================================================================

Private Const TBL_PEND As String = "Correção_Automática"

Public Sub AplicarFormulasPendencias()
    ' Mantido por compatibilidade (sem ações na arquitetura atual)
End Sub

' Mantida por compatibilidade, sem chamadas de sincronização (Opção A).
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

    ' Sem fórmulas na Correção_Automática (arquitetura atual)
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
