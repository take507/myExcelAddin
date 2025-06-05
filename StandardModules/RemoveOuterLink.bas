Sub 外部参照を置換(ws As Worksheet)

    Dim targetRange As Range
    Dim formulaArray As Variant
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim cellFormula As String

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

    
    ' ----------------------------------------------------
    ' 設定した範囲の数式をチェック
    ' ----------------------------------------------------
    Set targetRange = ws.UsedRange
    formulaArray = targetRange.Formula
    
    For r = 1 To UBound(formulaArray, 1)
        For c = 1 To UBound(formulaArray, 2)
            cellFormula = formulaArray(r, c)
            formulaArray(r, c) = 外部リンク文字列を置換(cellFormula)
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

Function 外部リンク文字列を置換(cellFormula As String)

    If cellFormula = "" Then
        外部リンク文字列を置換 = cellFormula
        Exit Function
    End If
    
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    ' =で始まる数式である
    ' C:path[filename.xlsx]sheetname! の形式を含む
    If Left(cellFormula, 1) = "=" And cellFormula Like "*:*[*xlsx]*!*" Then
        
        With reg
            .Pattern = "'[A-Z]:[^\]]+\]([^']+'!)"
            .IgnoreCase = False
            .Global = True
        End With
        
        Set Matches = reg.Execute(cellFormula)
        For Each Match In Matches
            cellFormula = Replace(cellFormula, Match.Value, "'" & Match.submatches(0))
        Next Match
    End If

    ' =で始まる数式である
    ' \\path[filename.xlsx]sheetname! の形式を含む
    If Left(cellFormula, 1) = "=" And cellFormula Like "*'\\*[*xlsx]*!*" Then
        With reg
            .Pattern = "'\\\\[^\]]+\]([^']+'!)"
            .IgnoreCase = False
            .Global = True
        End With
        
        Set Matches = reg.Execute(cellFormula)
        For Each Match In Matches
            cellFormula = Replace(cellFormula, Match.Value, "'" & Match.submatches(0))
        Next Match
    End If

    ' =で始まる数式である
    ' [filename.xlsx]sheetname! の形式を含む
    If Left(cellFormula, 1) = "=" And cellFormula Like "[*xlsx]*!*" Then
        
        With reg
            .Pattern = "'\[[^\]]+\]([^']+'!)"
            .IgnoreCase = False
            .Global = True
        End With
        
        Set Matches = reg.Execute(cellFormula)
        For Each Match In Matches
            cellFormula = Replace(cellFormula, Match.Value, "'" & Match.submatches(0))
        Next Match
    End If

    外部リンク文字列を置換 = cellFormula
    
End Function
