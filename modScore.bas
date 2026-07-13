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

'====================================================
' 未回答数
'====================================================
Public Function 未回答数() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    For r = 2 To LastRow

        If Nz(ws.Cells(r, scTotal).Value) = 0 Then

            未回答数 = 未回答数 + 1

        End If

    Next

End Function


'====================================================
' 全体正答率
'====================================================
Public Function 全体正答率() As Double

    If 総回答数() = 0 Then

        全体正答率 = 0

    Else

        全体正答率 = _
            総正解数() / 総回答数()

    End If

End Function

'====================================================
' 成績リセット
'====================================================
Public Sub 成績リセット()

    Dim ws As Worksheet
    Dim LastRow As Long

    If Not Confirm("成績をすべて削除しますか？") Then Exit Sub

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    If LastRow < 2 Then Exit Sub

    ws.Range( _
        ws.Cells(2, scTotal), _
        ws.Cells(LastRow, scLastDate) _
    ).ClearContents

    ShowMessage "成績をリセットしました。"

End Sub


'====================================================
' 問題の回答回数取得
'====================================================
Public Function 回答回数(ByVal ProblemID As Long) As Long

    Dim r As Long

    r = GetScoreRow(ProblemID)

    If r = 0 Then Exit Function

    回答回数 = _
        Nz(GetScoreWS.Cells(r, scTotal).Value)

End Function


'====================================================
' 問題の正答率取得
'====================================================
Public Function 問題正答率(ByVal ProblemID As Long) As Double

    Dim r As Long

    r = GetScoreRow(ProblemID)

    If r = 0 Then Exit Function

    問題正答率 = _
        Nz(GetScoreWS.Cells(r, scRate).Value)

End Function

'====================================================
' 回答済み問題数
'====================================================
Public Function 回答済み問題数() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    For r = 2 To LastRow

        If Nz(ws.Cells(r, scTotal).Value) > 0 Then

            回答済み問題数 = 回答済み問題数 + 1

        End If

    Next

End Function


'====================================================
' 学習進捗率
'====================================================
Public Function 学習進捗率() As Double

    Dim TotalProblem As Long

    TotalProblem = GetDataCount(GetMasterWS)

    If TotalProblem = 0 Then Exit Function

    学習進捗率 = 回答済み問題数 / TotalProblem

End Function


'====================================================
' 苦手問題か判定
'
' 正答率60%未満
'====================================================
Public Function IsWeakProblem(ByVal ProblemID As Long) As Boolean

    IsWeakProblem = _
        (問題正答率(ProblemID) < 0.6)

End Function


'====================================================
' 得意問題か判定
'
' 正答率90%以上
'====================================================
Public Function IsStrongProblem(ByVal ProblemID As Long) As Boolean

    IsStrongProblem = _
        (問題正答率(ProblemID) >= 0.9)

End Function

'====================================================
' 問題の正解数取得
'====================================================
Public Function 正解数(ByVal ProblemID As Long) As Long

    Dim r As Long

    r = GetScoreRow(ProblemID)

    If r = 0 Then Exit Function

    正解数 = Nz(GetScoreWS.Cells(r, scCorrect).Value)

End Function


'====================================================
' 問題の不正解数取得
'====================================================
Public Function 不正解数(ByVal ProblemID As Long) As Long

    Dim r As Long

    r = GetScoreRow(ProblemID)

    If r = 0 Then Exit Function

    不正解数 = Nz(GetScoreWS.Cells(r, scWrong).Value)

End Function

'====================================================
' 成績情報再計算
'
' 全問題の正答率を再計算
'====================================================
Public Sub 成績再計算()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = GetScoreWS()

    LastRow = GetLastRow(ws)

    For r = 2 To LastRow

        正答率更新 r

    Next

End Sub



