Attribute VB_Name = "modHistory"
Option Explicit

'====================================================
' 昇任試験学習システム
'
' Module : modHistory
' Version: 1.0.0
'
' 更新履歴
' 2026/07/12 Ver1.0.0 新規作成
'====================================================

'====================================================
' 履歴保存
'
' Result
'   ○
'   ×
'====================================================
Public Sub 履歴保存(ByVal Result As String)

    Dim wsH As Worksheet
    Dim wsM As Worksheet

    Dim RowNo As Long
    Dim NewRow As Long

    Set wsH = GetHistoryWS()
    Set wsM = GetMasterWS()

    RowNo = GetCurrentRow()

    If RowNo = 0 Then Exit Sub

    NewRow = GetLastRow(wsH) + 1

    wsH.Cells(NewRow, hcDate).Value = Now
    wsH.Cells(NewRow, hcID).Value = GetCurrentID()
    wsH.Cells(NewRow, hcUserAnswer).Value = _
        GetLearnWS.Range(CELL_USERANSWER).Value

    wsH.Cells(NewRow, hcCorrectAnswer).Value = _
        wsM.Cells(RowNo, mcAnswer).Value

    wsH.Cells(NewRow, hcResult).Value = Result

    wsH.Cells(NewRow, hcSubject).Value = _
        wsM.Cells(RowNo, mcSubject).Value

End Sub

'====================================================
' 履歴件数
'====================================================
Public Function 履歴件数() As Long

    履歴件数 = GetDataCount(GetHistoryWS)

End Function

'====================================================
' 履歴クリア
'====================================================
Public Sub 履歴クリア()

    Dim ws As Worksheet
    Dim LastRow As Long

    If Not Confirm("学習履歴を削除しますか？") Then Exit Sub

    Set ws = GetHistoryWS()

    LastRow = GetLastRow(ws)

    If LastRow < 2 Then Exit Sub

    ws.Range(ws.Rows(2), ws.Rows(LastRow)).ClearContents

    ShowMessage "学習履歴を削除しました。"

End Sub

'====================================================
' 最新履歴行取得
'====================================================
Private Function 最新履歴行() As Long

    最新履歴行 = GetLastRow(GetHistoryWS)

End Function

