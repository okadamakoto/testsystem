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

'====================================================
' 回答正規化
'
' ・全角スペース除去
' ・半角スペース除去
' ・タブ除去
' ・改行除去
' ・①②③④を昇順へ並べ替え
'====================================================
Private Function NormalizeAnswer(ByVal Answer As String) As String

    Answer = RemoveSpaces(Answer)

    Answer = SortAnswer(Answer)

    NormalizeAnswer = Answer

End Function


'====================================================
' 空白除去
'====================================================
Private Function RemoveSpaces(ByVal Text As String) As String

    Text = Replace(Text, " ", "")
    Text = Replace(Text, "　", "")
    Text = Replace(Text, vbTab, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, vbLf, "")

    RemoveSpaces = Trim(Text)

End Function


'====================================================
' 回答並べ替え
'
' ④①
' ↓
' ①④
'====================================================
Private Function SortAnswer(ByVal Answer As String) As String

    Dim i As Long
    Dim Result As String

    Result = ""

    For i = 1 To 4

        If InStr(Answer, GetChoice(i)) > 0 Then

            Result = Result & GetChoice(i)

        End If

    Next i

    SortAnswer = Result

End Function


'====================================================
' 選択肢文字取得
'====================================================
Private Function GetChoice(ByVal No As Long) As String

    Select Case No

        Case 1
            GetChoice = "①"

        Case 2
            GetChoice = "②"

        Case 3
            GetChoice = "③"

        Case 4
            GetChoice = "④"

        Case Else
            GetChoice = ""

    End Select

End Function


'====================================================
' 正誤判定
'====================================================
Private Function IsCorrect( _
            ByVal UserAnswer As String, _
            ByVal CorrectAnswer As String) As Boolean

    IsCorrect = (UserAnswer = CorrectAnswer)

End Function

