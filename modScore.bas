Attribute VB_Name = "modScore"
Option Explicit

'====================================================
' 昇任試験学習システム
'
' Module : modScore
' Version: 1.0.0
'
' 更新履歴
' 2026/07/12 Ver1.0.0 新規作成
'====================================================

'====================================================
' 成績更新
'
' IsCorrect
'   True  ：正解
'   False ：不正解
'====================================================
Public Sub 成績更新(ByVal IsCorrect As Boolean)

    Dim ws As Worksheet

    Dim ProblemID As Long
    Dim ScoreRow As Long

    Set ws = GetScoreWS()

    ProblemID = GetCurrentID()

    If ProblemID = 0 Then Exit Sub

    ScoreRow = GetScoreRow(ProblemID)

    If ScoreRow = 0 Then Exit Sub

    '回答回数
    ws.Cells(ScoreRow, scTotal).Value = _
        Nz(ws.Cells(ScoreRow, scTotal).Value) + 1

    If IsCorrect Then

        '正解数
        ws.Cells(ScoreRow, scCorrect).Value = _
            Nz(ws.Cells(ScoreRow, scCorrect).Value) + 1

    Else

        '不正解数
        ws.Cells(ScoreRow, scWrong).Value = _
            Nz(ws.Cells(ScoreRow, scWrong).Value) + 1

    End If

    '最終回答日
    ws.Cells(ScoreRow, scLastDate).Value = Now

    Call 正答率更新(ScoreRow)

End Sub

'====================================================
' 正答率更新
'====================================================
Private Sub 正答率更新(ByVal ScoreRow As Long)

    Dim ws As Worksheet

    Dim Total As Long
    Dim Correct As Long

    Set ws = GetScoreWS()

    Total = Nz(ws.Cells(ScoreRow, scTotal).Value)
    Correct = Nz(ws.Cells(ScoreRow, scCorrect).Value)

    If Total = 0 Then

        ws.Cells(ScoreRow, scRate).Value = 0

    Else

        ws.Cells(ScoreRow, scRate).Value = Correct / Total

    End If

End Sub

