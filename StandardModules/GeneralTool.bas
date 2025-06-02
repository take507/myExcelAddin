Sub �Ԙg�ǉ�()
    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left, Selection.Top, Selection.Width, Selection.Height)
        ' �h��Ԃ�����
        .Fill.Visible = msoFalse
        ' �r���ݒ�
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Transparency = 0
        .Line.Weight = 3
    End With
End Sub
Sub �Ԗ��ǉ�()
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
Sub �t�@�C���̏ꏊ���J��()
    Shell "explorer.exe /select," & ActiveWorkbook.FullName, vbNormalFocus
End Sub
Sub �t�@�C���p�X���N���b�v�{�[�h�ێ�()
    ClipBoardSave (ActiveWorkbook.FullName)
End Sub
Sub �O���Q�Ƃ�u��()

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim formulaArray As Variant
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim cellFormula As String

    Set ws = Sheets("Sheet1")

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

    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "'C:[^\]]+\]([^']+'!)"
        .IgnoreCase = False
        .Global = True
    End With
    
    ' ----------------------------------------------------
    ' �ݒ肵���͈͂̐������`�F�b�N
    ' ----------------------------------------------------
    Set targetRange = ws.UsedRange
    formulaArray = targetRange.Formula
    
    For r = 1 To UBound(formulaArray, 1)
        For c = 1 To UBound(formulaArray, 2)
            cellFormula = formulaArray(r, c)
            ' ��ł͂Ȃ�
            ' =�Ŏn�܂鐔���ł���
            ' C:path[filename.xlsx]sheetname! �̌`�����܂�
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
    ' ������ɐݒ�����ɖ߂�
    ' ----------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
