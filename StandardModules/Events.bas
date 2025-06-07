Sub シートの整理(book As Workbook)

    Application.ScreenUpdating = False

    ' シートの件数分ループ
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        sheet.Activate
        
        ' シートの表示倍率を100%に戻す
        ActiveWindow.Zoom = 100
        ' ウィンドウの固定がされている場合
        If ActiveWindow.FreezePanes = True Then
            ' 浮動領域の左上を選択
            sheet.Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select
        Else
            ' A1を選択
            sheet.Range("A1").Select
        End If
        ' スクロール位置をセルの場所に修正
        ActiveWindow.ScrollRow = ActiveCell.Row
        ActiveWindow.ScrollColumn = ActiveCell.Column
    
    Next
    ' 先頭のシートを選択した状態にする
    book.Worksheets(1).Activate

    Call 無効な名前の定義の削除
    Call スタイルフォント調整

    Application.ScreenUpdating = True
    Exit Sub
          
End Sub

