Attribute VB_Name = "modQuestion"
Option Explicit

'====================================================
' 昇任試験学習システム
'
' Module : modMain
' Version: 1.0.0
'
' 更新履歴
' 2026/07/11 Ver1.0.0 新規作成
'====================================================

'====================================================
' 問題表示
'
' RowNo：問題マスタの行番号
'====================================================
Public Sub 表示問題(ByVal RowNo As Long)

    Dim wsM As Worksheet
    Dim wsL As Worksheet

    If Not IsValidRow(RowNo) Then
        ShowError MSG_NOTFOUND
        Exit Sub
    End If

    Set wsM = GetMasterWS()
    Set wsL = GetLearnWS()

    With wsL

        .Range(CELL_ID).Value = _
            wsM.Cells(RowNo, COL_ID).Value

        .Range(CELL_CATEGORY).Value = _
            wsM.Cells(RowNo, COL_CATEGORY).Value

        .Range(CELL_NUMBER).Value = _
            wsM.Cells(RowNo, COL_NUMBER).Value

        .Range(CELL_SUBJECT).Value = _
            wsM.Cells(RowNo, COL_SUBJECT).Value

        .Range(CELL_SOURCE).Value = _
            wsM.Cells(RowNo, COL_SOURCE).Value

        .Range(CELL_QUESTION).Value = _
            wsM.Cells(RowNo, COL_QUESTION).Value

    End With

    ClearAnswer

End Sub

'====================================================
' 最初の問題
'====================================================
Public Sub 最初の問題()

    Dim LastRow As Long

    LastRow = GetLastRow(GetMasterWS)

    If LastRow < 2 Then
        ShowError MSG_NO_DATA
        Exit Sub
    End If

    表示問題 2

End Sub


'====================================================
' 最後の問題
'====================================================
Public Sub 最後の問題()

    Dim LastRow As Long

    LastRow = GetLastRow(GetMasterWS)

    If LastRow < 2 Then
        ShowError MSG_NO_DATA
        Exit Sub
    End If

    表示問題 LastRow

End Sub


'====================================================
' 次の問題
'====================================================
Public Sub 次の問題()

    Dim RowNo As Long
    Dim LastRow As Long

    RowNo = GetCurrentRow()

    If RowNo = 0 Then

        最初の問題

        Exit Sub

    End If

    LastRow = GetLastRow(GetMasterWS)

    If RowNo >= LastRow Then

        ShowMessage MSG_LAST

        Exit Sub

    End If

    表示問題 RowNo + 1

End Sub


'====================================================
' 前の問題
'====================================================
Public Sub 前の問題()

    Dim RowNo As Long

    RowNo = GetCurrentRow()

    If RowNo = 0 Then

        最初の問題

        Exit Sub

    End If

    If RowNo <= 2 Then

        ShowMessage MSG_FIRST

        Exit Sub

    End If

    表示問題 RowNo - 1

End Sub

'====================================================
' ランダム出題
'====================================================
Public Sub ランダム出題()

    Dim LastRow As Long
    Dim RowNo As Long

    LastRow = GetLastRow(GetMasterWS)

    If LastRow < 2 Then
        ShowError MSG_NO_DATA
        Exit Sub
    End If

    Randomize

    RowNo = Int((LastRow - 1) * Rnd) + 2

    表示問題 RowNo

End Sub

'====================================================
' 指定問題へ移動
'
' B2へ入力された連番へ移動
'====================================================
Public Sub 指定問題へ移動()

    Dim ID As Long
    Dim RowNo As Long

    With GetLearnWS

        If Trim(.Range(CELL_ID).Value) = "" Then

            ShowError MSG_INPUTID
            Exit Sub

        End If

        If Not IsNumeric(.Range(CELL_ID).Value) Then

            ShowError MSG_NUMERIC
            Exit Sub

        End If

        ID = CLng(.Range(CELL_ID).Value)

    End With

    RowNo = GetRowByID(ID)

    If RowNo = 0 Then

        ShowError MSG_NOTFOUND
        Exit Sub

    End If

    表示問題 RowNo

End Sub

'====================================================
' 現在の問題を再表示
'====================================================
Public Sub 現在の問題を再表示()

    Dim RowNo As Long

    RowNo = GetCurrentRow()

    If RowNo = 0 Then

        ShowError MSG_NOTFOUND
        Exit Sub

    End If

    表示問題 RowNo

End Sub

'====================================================
' 問題表示初期化
'====================================================
Public Sub 問題表示初期化()

    ClearAnswer

    With GetLearnWS

        .Range(CELL_ID).ClearContents
        .Range(CELL_CATEGORY).ClearContents
        .Range(CELL_NUMBER).ClearContents
        .Range(CELL_SUBJECT).ClearContents
        .Range(CELL_SOURCE).ClearContents
        .Range(CELL_QUESTION).ClearContents

    End With

End Sub

