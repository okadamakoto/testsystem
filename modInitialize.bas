Attribute VB_Name = "modInitialize"
Option Explicit

'====================================================
' 昇任試験学習システム
'
' Module : modInitialize
' Version: 1.0.0
'
' 更新履歴
' 2026/07/12 Ver1.0.0 新規作成
'====================================================

'====================================================
' 初期化
'====================================================
Public Sub 初期化()

    Application.ScreenUpdating = False

    初期化_成績シート

    初期化_学習履歴

    問題表示初期化

    最初の問題

    Application.ScreenUpdating = True

    ShowMessage "初期化が完了しました。"

End Sub


'====================================================
' 成績シート初期化
'====================================================
Private Sub 初期化_成績シート()

    Dim wsM As Worksheet
    Dim wsS As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Set wsM = GetMasterWS()
    Set wsS = GetScoreWS()

    wsS.Cells.ClearContents

    wsS.Range("A1") = "連番"
    wsS.Range("B1") = "回答回数"
    wsS.Range("C1") = "正解数"
    wsS.Range("D1") = "不正解数"
    wsS.Range("E1") = "正答率"
    wsS.Range("F1") = "最終回答日"

    LastRow = GetLastRow(wsM)

    For r = 2 To LastRow

        wsS.Cells(r, 1) = wsM.Cells(r, COL_ID)

        wsS.Cells(r, 2) = 0
        wsS.Cells(r, 3) = 0
        wsS.Cells(r, 4) = 0
        wsS.Cells(r, 5) = 0
        wsS.Cells(r, 6) = ""

    Next

End Sub


'====================================================
' 学習履歴初期化
'====================================================
Private Sub 初期化_学習履歴()

    Dim ws As Worksheet

    Set ws = GetHistoryWS()

    ws.Cells.ClearContents

    ws.Range("A1") = "日時"
    ws.Range("B1") = "連番"
    ws.Range("C1") = "回答"
    ws.Range("D1") = "正答"
    ws.Range("E1") = "判定"
    ws.Range("F1") = "科目"

End Sub


'====================================================
' メニュー初期化
'====================================================
Public Sub メニュー初期化()

    問題表示初期化

End Sub


'====================================================
' 問題マスタ更新
'====================================================
Public Sub 問題マスタ更新()

    初期化_成績シート

    ShowMessage "成績シートを更新しました。"

End Sub

