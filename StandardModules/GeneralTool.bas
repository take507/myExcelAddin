Sub 赤枠追加()
    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left, Selection.Top, Selection.Width, Selection.Height)
        ' 塗りつぶし無し
        .Fill.Visible = msoFalse
        ' 罫線設定
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Transparency = 0
        .Line.Weight = 3
    End With
End Sub
Sub 赤矢印追加()
    With ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Selection.Left, Selection.Top, Selection.Left + Selection.Width, Selection.Top + Selection.Height)
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.BeginArrowheadStyle = msoArrowheadNone
        .Line.EndArrowheadStyle = msoArrowheadOpen
        .Line.Visible = msoTrue
        .Line.Weight = 2.75
        .Line.EndArrowheadLength = msoArrowheadLengthMedium
        .Line.EndArrowheadWidth = msoArrowheadWide
    End With
End Sub
Sub ファイルの場所を開く()
    Shell "explorer.exe /select," & ActiveWorkbook.FullName, vbNormalFocus
End Sub
Sub ファイルパスをクリップボード保持()
    ClipBoardSave (ActiveWorkbook.FullName)
End Sub
Sub 無効な名前の定義の削除()
    
    ' 名前の定義の件数分ループ
    Dim name As Object
    For Each name In Names
        ' 表示にする
        name.Visible = True
        ' 参照先がエラーになっている
        If InStr(name.RefersTo, "#REF") > 0 Then
            Debug.Print ("無効な名前の定義の削除[name:" & name.name & ",RefersTo" & name.RefersTo & "]")
            ' 削除する
            name.Delete
        End If
    Next

End Sub
Sub スタイルフォント調整()
    Dim normalStyle As Style
    Set normalStyle = ActiveWorkbook.Styles("Normal")
    normalStyle.Font.name = Range("A1").Font.name
    
    For Each vStyle In ActiveWorkbook.Styles
        If (vStyle.name = "Hyperlink" Or vStyle.name = "Followed Hyperlink") And vStyle.BuiltIn = True Then
            vStyle.Font.name = normalStyle.Font.name
            vStyle.Font.Size = normalStyle.Font.Size
        End If
    Next
End Sub
Sub シート一覧()

    For Each sheet In Worksheets
        If sheet.Visible = True Then
            Debug.Print sheet.name
        End If
    Next

End Sub
