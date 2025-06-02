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
Sub 外部参照を置換()

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim formulaArray As Variant
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim cellFormula As String

    Set ws = Sheets("Sheet1")

    ' ----------------------------------------------------
    ' パフォーマンス向上のための設定
    ' ----------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ----------------------------------------------------
    ' 処理対象の動的な範囲を設定
    ' ----------------------------------------------------
    If ws.UsedRange.Cells.count = 1 And IsEmpty(ws.UsedRange.Cells(1, 1).Value) Then
        GoTo CleanUp
    End If

    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "'C:[^\]]+\]([^']+'!)"
        .IgnoreCase = False
        .Global = True
    End With
    
    ' ----------------------------------------------------
    ' 設定した範囲の数式をチェック
    ' ----------------------------------------------------
    Set targetRange = ws.UsedRange
    formulaArray = targetRange.Formula
    
    For r = 1 To UBound(formulaArray, 1)
        For c = 1 To UBound(formulaArray, 2)
            cellFormula = formulaArray(r, c)
            ' 空ではない
            ' =で始まる数式である
            ' C:path[filename.xlsx]sheetname! の形式を含む
            If cellFormula <> "" And Left(cellFormula, 1) = "=" And cellFormula Like "*C:*[*xlsx]*!*" Then
                Set Matches = reg.Execute(cellFormula)
                For Each Match In Matches
                    formulaArray(r, c) = Replace(cellFormula, Match.Value, "'" & Match.submatches(0))
                Next Match
            End If
        Next c
    Next r
    targetRange.Formula = formulaArray

CleanUp:
    ' ----------------------------------------------------
    ' 処理後に設定を元に戻す
    ' ----------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
