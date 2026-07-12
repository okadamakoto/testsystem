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


'====================================================
' 成績シートの行番号取得
'====================================================
Private Function GetScoreRow(ByVal ProblemID As Long) As Long

    Dim r As Variant

    r = Application.Match(ProblemID, _
                          GetScoreWS.Columns(scID), _
                          0)

    If IsError(r) Then
        GetScoreRow = 0
    Else
        GetScoreRow = CLng(r)
    End If

End Function


'====================================================
' 総回答数
'====================================================
Public Function 総回答数() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    For r = 2 To LastRow

        総回答数 = _
            総回答数 + Nz(ws.Cells(r, scTotal).Value)

    Next

End Function


'====================================================
' 総正解数
'====================================================
Public Function 総正解数() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    For r = 2 To LastRow

        総正解数 = _
            総正解数 + Nz(ws.Cells(r, scCorrect).Value)

    Next

End Function


'====================================================
' 総不正解数
'====================================================
Public Function 総不正解数() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    For r = 2 To LastRow

        総不正解数 = _
            総不正解数 + Nz(ws.Cells(r, scWrong).Value)

    Next

End Function


