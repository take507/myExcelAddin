Sub �V�[�g�̐���(book As Workbook)

    Application.ScreenUpdating = False

    ' �V�[�g�̌��������[�v
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        sheet.Activate
        
        ' �V�[�g�̕\���{����100%�ɖ߂�
        ActiveWindow.Zoom = 100
        ' �E�B���h�E�̌Œ肪����Ă���ꍇ
        If ActiveWindow.FreezePanes = True Then
            ' �����̈�̍����I��
            sheet.Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select
        Else
            ' A1��I��
            sheet.Range("A1").Select
        End If
        ' �X�N���[���ʒu���Z���̏ꏊ�ɏC��
        ActiveWindow.ScrollRow = ActiveCell.Row
        ActiveWindow.ScrollColumn = ActiveCell.Column
    
    Next
    ' �擪�̃V�[�g��I��������Ԃɂ���
    book.Worksheets(1).Activate

    Call �����Ȗ��O�̒�`�̍폜
    Call �X�^�C���t�H���g����

    Application.ScreenUpdating = True
    Exit Sub
          
End Sub

