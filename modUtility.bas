Attribute VB_Name = "modUtility"
Option Explicit

'====================================================
' 昇任試験学習システム Ver1.0
' 共通関数
'====================================================

'====================================================
' シート取得
'====================================================

Public Function GetMasterWS() As Worksheet

    Set GetMasterWS = ThisWorkbook.Worksheets(SHEET_MASTER)

End Function

Public Function GetLearnWS() As Worksheet

    Set GetLearnWS = ThisWorkbook.Worksheets(SHEET_LEARN)

End Function

Public Function GetScoreWS() As Worksheet

    Set GetScoreWS = ThisWorkbook.Worksheets(SHEET_SCORE)

End Function

Public Function GetHistoryWS() As Worksheet

    Set GetHistoryWS = ThisWorkbook.Worksheets(SHEET_HISTORY)

End Function

Public Function GetSettingWS() As Worksheet

    Set GetSettingWS = ThisWorkbook.Worksheets(SHEET_SETTING)

End Function

Public Function GetDashboardWS() As Worksheet

    Set GetDashboardWS = ThisWorkbook.Worksheets(SHEET_DASHBOARD)

End Function

'====================================================
' 最終行取得
'====================================================
Public Function GetLastRow( _
            ByVal ws As Worksheet, _
            Optional ByVal Col As Long = COL_ID) As Long

    If ws Is Nothing Then Exit Function

    GetLastRow = ws.Cells(ws.Rows.Count, Col).End(xlUp).Row

End Function

'====================================================
' 現在表示中の問題ID
'====================================================
Public Function GetCurrentID() As Long

    With GetLearnWS()

        If IsNumeric(.Range(CELL_ID).Value) Then

            GetCurrentID = CLng(.Range(CELL_ID).Value)

        End If

    End With

End Function

'====================================================
' 問題ID→行番号
'====================================================
Public Function GetRowByID(ByVal ID As Long) As Long

    Dim r As Variant

    If ID <= 0 Then Exit Function

    r = Application.Match(ID, GetMasterWS.Columns(COL_ID), 0)

    If Not IsError(r) Then

        GetRowByID = CLng(r)

    End If

End Function

'====================================================
' 現在の行番号
'====================================================
Public Function GetCurrentRow() As Long

    GetCurrentRow = GetRowByID(GetCurrentID)

End Function

'====================================================
' 問題存在確認
'====================================================
Public Function ExistsProblem(ByVal ID As Long) As Boolean

    ExistsProblem = (GetRowByID(ID) > 0)

End Function

'====================================================
' 行番号チェック
'====================================================
Public Function IsValidRow(ByVal RowNo As Long) As Boolean

    IsValidRow = _
        (RowNo >= 2 And _
        RowNo <= GetLastRow(GetMasterWS))

End Function

'====================================================
' 回答欄クリア
'====================================================
Public Sub ClearAnswer()

    With GetLearnWS()

        .Range(CELL_USERANSWER).ClearContents
        .Range(CELL_RESULT).ClearContents
        .Range(CELL_EXPLANATION).ClearContents

    End With

End Sub

'====================================================
' メッセージ
'====================================================
Public Sub ShowMessage(ByVal Msg As String)

    MsgBox Msg, vbInformation, APP_TITLE

End Sub

'====================================================
' エラー
'====================================================
Public Sub ShowError(ByVal Msg As String)

    MsgBox Msg, vbExclamation, APP_TITLE

End Sub

'====================================================
' 確認
'====================================================
Public Function Confirm(ByVal Msg As String) As Boolean

    Confirm = _
        (MsgBox(Msg, _
        vbYesNo + vbQuestion, _
        APP_TITLE) = vbYes)

End Function


