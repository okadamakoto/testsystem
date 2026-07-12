Attribute VB_Name = "modAnswer"
Option Explicit

'====================================================
' 昇任試験学習システム
'
' Module : modAnswer
' Version: 1.0.0
'
' 更新履歴
' 2026/07/12 Ver1.0.0 新規作成
'====================================================

'====================================================
' 答え合わせ
'====================================================
Public Sub 答え合わせ()

    Dim wsM As Worksheet
    Dim wsL As Worksheet

    Dim RowNo As Long

    Dim UserAnswer As String
    Dim CorrectAnswer As String

    Set wsM = GetMasterWS()
    Set wsL = GetLearnWS()

    RowNo = GetCurrentRow()

    If RowNo = 0 Then

        ShowError "問題が表示されていません。"

        Exit Sub

    End If

    UserAnswer = Trim(wsL.Range(CELL_USERANSWER).Value)

    If UserAnswer = "" Then

        ShowError "回答を入力してください。"

        Exit Sub

    End If

    CorrectAnswer = wsM.Cells(RowNo, mcAnswer).Value

    UserAnswer = NormalizeAnswer(UserAnswer)

    CorrectAnswer = NormalizeAnswer(CorrectAnswer)

    If IsCorrect(UserAnswer, CorrectAnswer) Then

        wsL.Range(CELL_RESULT).Value = RESULT_CORRECT

        成績更新 True

        履歴保存 RESULT_OK

    Else

        wsL.Range(CELL_RESULT).Value = RESULT_WRONG

        成績更新 False

        履歴保存 RESULT_NG

    End If

End Sub

'====================================================
' 解説表示
'====================================================
Public Sub 解説表示()

    Dim wsM As Worksheet
    Dim wsL As Worksheet

    Dim RowNo As Long

    Set wsM = GetMasterWS()
    Set wsL = GetLearnWS()

    RowNo = GetCurrentRow()

    If RowNo = 0 Then Exit Sub

    wsL.Range(CELL_EXPLANATION).Value = _
            "【正答】" & _
            wsM.Cells(RowNo, mcAnswer).Value & _
            vbCrLf & vbCrLf & _
            wsM.Cells(RowNo, mcExplanation).Value

End Sub
