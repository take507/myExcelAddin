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
