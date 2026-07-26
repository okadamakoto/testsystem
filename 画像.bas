Option Explicit

Sub TestDisplayImage()

    Dim wsImage As Worksheet
    Dim wsDisplay As Worksheet

    Dim shp As Shape
    Dim targetCell As Range

    '画像を保存しているシート
    Set wsImage = ThisWorkbook.Worksheets("Sheet1")

    '問題を表示するシート
    Set wsDisplay = ThisWorkbook.Worksheets("Sheet2")

    '表示先
    Set targetCell = wsDisplay.Range("A20")

    '元画像を取得
    Set shp = wsImage.Shapes("ST_R07_AM1_Q2")

    '画像をコピー
    shp.Copy

    '貼り付け
    wsDisplay.Activate
    wsDisplay.Paste

    '貼り付けた画像をA20に移動
    With Selection
        .Left = targetCell.Left
        .Top = targetCell.Top
    End With

End Sub
