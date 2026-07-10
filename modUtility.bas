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
