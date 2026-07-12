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

'====================================================
' 回答入力チェック
'====================================================
Private Function ValidateAnswer(ByVal UserAnswer As String) As Boolean

    Dim i As Long
    Dim ch As String

    UserAnswer = RemoveSpaces(UserAnswer)

    If UserAnswer = "" Then

        ShowError "回答を入力してください。"

        Exit Function

    End If

    For i = 1 To Len(UserAnswer)

        ch = Mid$(UserAnswer, i, 1)

        Select Case ch

            Case "①", "②", "③", "④"

            Case Else

                ShowError "回答は①②③④のみ入力できます。"

                Exit Function

        End Select

    Next i

    ValidateAnswer = True

End Function


'====================================================
' 重複回答チェック
'====================================================
Private Function HasDuplicateAnswer(ByVal UserAnswer As String) As Boolean

    UserAnswer = NormalizeAnswer(UserAnswer)

    If Len(UserAnswer) <> Len(RemoveDuplicate(UserAnswer)) Then

        HasDuplicateAnswer = True

    End If

End Function


'====================================================
' 重複除去
'====================================================
Private Function RemoveDuplicate(ByVal Text As String) As String

    Dim Result As String
    Dim i As Long
    Dim ch As String

    Result = ""

    For i = 1 To Len(Text)

        ch = Mid$(Text, i, 1)

        If InStr(Result, ch) = 0 Then

            Result = Result & ch

        End If

    Next i

    RemoveDuplicate = Result

End Function


'====================================================
' 回答チェック
'====================================================
Public Function 回答チェック() As Boolean

    Dim Ans As String

    Ans = GetLearnWS.Range(CELL_USERANSWER).Value

    If Not ValidateAnswer(Ans) Then Exit Function

    If HasDuplicateAnswer(Ans) Then

        ShowError "同じ選択肢が重複しています。"

        Exit Function

    End If

    回答チェック = True

End Function


'====================================================
' 解説クリア
'====================================================
Public Sub 解説クリア()

    GetLearnWS.Range(CELL_EXPLANATION).ClearContents

End Sub


'====================================================
' 判定クリア
'====================================================
Public Sub 判定クリア()

    GetLearnWS.Range(CELL_RESULT).ClearContents

End Sub


'====================================================
' 回答クリア
'====================================================
Public Sub 回答クリア()

    GetLearnWS.Range(CELL_USERANSWER).ClearContents

End Sub


'====================================================
' 回答関連クリア
'====================================================
Public Sub 回答初期化()

    回答クリア

    判定クリア

    解説クリア

End Sub


'====================================================
' 解説表示切替
'====================================================
Public Sub 解説表示切替()

    With GetLearnWS.Range(CELL_EXPLANATION)

        If Trim(.Value) = "" Then

            解説表示

        Else

            .ClearContents

        End If

    End With

End Sub


