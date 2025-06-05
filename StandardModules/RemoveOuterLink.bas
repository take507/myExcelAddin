Sub �O���Q�Ƃ�u��(ws As Worksheet)

    Dim targetRange As Range
    Dim formulaArray As Variant
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim cellFormula As String

    ' ----------------------------------------------------
    ' �p�t�H�[�}���X����̂��߂̐ݒ�
    ' ----------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ----------------------------------------------------
    ' �����Ώۂ̓��I�Ȕ͈͂�ݒ�
    ' ----------------------------------------------------
    If ws.UsedRange.Cells.count = 1 And IsEmpty(ws.UsedRange.Cells(1, 1).Value) Then
        GoTo CleanUp
    End If

    
    ' ----------------------------------------------------
    ' �ݒ肵���͈͂̐������`�F�b�N
    ' ----------------------------------------------------
    Set targetRange = ws.UsedRange
    formulaArray = targetRange.Formula
    
    For r = 1 To UBound(formulaArray, 1)
        For c = 1 To UBound(formulaArray, 2)
            cellFormula = formulaArray(r, c)
            formulaArray(r, c) = �O�������N�������u��(cellFormula)
        Next c
    Next r
    targetRange.Formula = formulaArray

CleanUp:
    ' ----------------------------------------------------
    ' ������ɐݒ�����ɖ߂�
    ' ----------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Function �O�������N�������u��(cellFormula As String)

    If cellFormula = "" Then
        �O�������N�������u�� = cellFormula
        Exit Function
    End If
    
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    ' =�Ŏn�܂鐔���ł���
    ' C:path[filename.xlsx]sheetname! �̌`�����܂�
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

    ' =�Ŏn�܂鐔���ł���
    ' \\path[filename.xlsx]sheetname! �̌`�����܂�
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

    ' =�Ŏn�܂鐔���ł���
    ' [filename.xlsx]sheetname! �̌`�����܂�
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

    �O�������N�������u�� = cellFormula
    
End Function
