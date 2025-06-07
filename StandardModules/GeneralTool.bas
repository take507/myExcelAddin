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
Sub �����Ȗ��O�̒�`�̍폜()
    
    ' ���O�̒�`�̌��������[�v
    Dim name As Object
    For Each name In Names
        ' �\���ɂ���
        name.Visible = True
        ' �Q�Ɛ悪�G���[�ɂȂ��Ă���
        If InStr(name.RefersTo, "#REF") > 0 Then
            Debug.Print ("�����Ȗ��O�̒�`�̍폜[name:" & name.name & ",RefersTo" & name.RefersTo & "]")
            ' �폜����
            name.Delete
        End If
    Next

End Sub
Sub �X�^�C���t�H���g����()
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
Sub �V�[�g�ꗗ()

    For Each sheet In Worksheets
        If sheet.Visible = True Then
            Debug.Print sheet.name
        End If
    Next

End Sub
