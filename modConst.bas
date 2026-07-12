Attribute VB_Name = "modConst"
Option Explicit

'====================================================
' 昇任試験学習システム
'
' Module : modConst
' Version: 1.0.0
'
' 更新履歴
' 2026/07/12 Ver1.0.0 新規作成
'====================================================

'====================================================
' シート名
'====================================================

Public Const SHEET_MASTER      As String = "問題マスタ"
Public Const SHEET_LEARN       As String = "学習"
Public Const SHEET_SCORE       As String = "成績"
Public Const SHEET_HISTORY     As String = "学習履歴"
Public Const SHEET_SETTING     As String = "設定"
Public Const SHEET_DASHBOARD   As String = "ダッシュボード"

'====================================================
' 問題マスタ列番号
'
' A:連番
' B:管理番号（区分）
' C:管理番号（No.）
' D:科目
' E:出題元
' F:問題
' G:正答
' H:解説
'====================================================

Public Enum MasterColumn

    mcID = 1
    mcCategory = 2
    mcNumber = 3
    mcSubject = 4
    mcSource = 5
    mcQuestion = 6
    mcAnswer = 7
    mcExplanation = 8

End Enum

'====================================================
' 学習シート
'====================================================

Public Const CELL_ID           As String = "B2"
Public Const CELL_CATEGORY     As String = "B3"
Public Const CELL_NUMBER       As String = "B4"
Public Const CELL_SUBJECT      As String = "B5"
Public Const CELL_SOURCE       As String = "B6"
Public Const CELL_QUESTION     As String = "B7"

Public Const CELL_USERANSWER   As String = "B20"
Public Const CELL_RESULT       As String = "B22"
Public Const CELL_EXPLANATION  As String = "B24"

'====================================================
' 成績シート列番号
'====================================================

Public Enum ScoreColumn

    scID = 1
    scTotal = 2
    scCorrect = 3
    scWrong = 4
    scRate = 5
    scLastDate = 6

End Enum

'====================================================
' 学習履歴列番号
'====================================================

Public Enum HistoryColumn

    hcDate = 1
    hcID = 2
    hcUserAnswer = 3
    hcCorrectAnswer = 4
    hcResult = 5
    hcSubject = 6

End Enum

'====================================================
' アプリケーション
'====================================================

Public Const APP_TITLE As String = "昇任試験学習システム"

'====================================================
' メッセージ
'====================================================

Public Const MSG_NO_DATA      As String = "問題が登録されていません。"

Public Const MSG_FIRST        As String = "最初の問題です。"

Public Const MSG_LAST         As String = "最後の問題です。"

Public Const MSG_NOTFOUND     As String = "指定された問題は存在しません。"

Public Const MSG_INPUTID      As String = "連番を入力してください。"

Public Const MSG_NUMERIC      As String = "連番は数値で入力してください。"

Public Const MSG_INIT_END     As String = "初期化が完了しました。"

Public Const MSG_SCORE_UPDATE As String = "成績シートを更新しました。"

'====================================================
' 判定
'====================================================

Public Const RESULT_OK        As String = "○"
Public Const RESULT_NG        As String = "×"

Public Const RESULT_CORRECT   As String = "○ 正解"
Public Const RESULT_WRONG     As String = "× 不正解"

'====================================================
' 回答区切り文字
'====================================================

Public Const ANSWER_SEPARATOR As String = " "

'====================================================
' バージョン情報
'====================================================

Public Const APP_VERSION As String = "1.0.0"
