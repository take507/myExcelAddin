Sub シート一覧()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = Sheets.Add

    ws.name = "シート一覧"

    ' シート全体のフォントを設定
    With ws.Cells.Font
        .name = "Meiryo UI"
        .Size = 9
    End With

    ' 1行目のセルに項目名を設定
    ws.Cells(1, 1).Value = "No"
    ws.Cells(1, 2).Value = "シート名"
    ws.Cells(1, 3).Value = "備考"
    
    ' テーブルの作成とスタイル適用
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C2"), , xlYes)
    ' テーブルスタイルを適用
    lo.TableStyle = "TableStyleMedium2"
    ' オートフィルタ設定
    lo.ShowAutoFilter = True
    ' 縞模様
    lo.ShowTableStyleRowStripes = True
    ' 合計行
    lo.ShowTotals = False
    ' ヘッダ行背景色
    lo.HeaderRowRange.Interior.Color = RGB(64, 84, 106)
    ' 最初の列に特別な書式を適用する
    lo.ShowTableStyleFirstColumn = False
    ' 最後の列に特別な書式を適用する
    lo.ShowTableStyleLastColumn = False

    ' ウィンドウ枠をB2セルで固定
    With Application.ActiveWindow
        .FreezePanes = False
        ws.Activate
        ws.Range("B2").Select
        .FreezePanes = True
    End With

    With ws.Columns("A:C")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.Color = RGB(192, 192, 192)
    End With

    ws.Columns("D:XFD").Hidden = True
    ws.Range("A2").Formula = "=ROW()-1"
    
    ' シート名出力
    Dim row As Integer
    row = 2
    For Each sheet In Worksheets
        If sheet.Visible = True Then
            ws.Cells(row, 2).Value = sheet.name
            row = row + 1
        End If
    Next
    
    ' 列幅を自動調整
    ws.Columns("A:H").AutoFit
    ws.Columns("C").ColumnWidth = 40

End Sub
