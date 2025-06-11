Sub INSERT�������V�[�g()

    Const SHEET_NAME_SQLINSERT As String = "INSERT������"

    If ExistsSheet(SHEET_NAME_SQLINSERT) = False Then
        ThisWorkbook.Sheets(SHEET_NAME_SQLINSERT).Copy After:=Worksheets(Worksheets.Count)
    Else
        MsgBox (SHEET_NAME_SQLINSERT & "�V�[�g�����ɑ��݂��Ă��܂��B")
    End If

End Sub

Sub �V�[�g�ꗗ()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = Sheets.Add

    ws.name = "�V�[�g�ꗗ"

    ' �V�[�g�S�̂̃t�H���g��ݒ�
    With ws.Cells.Font
        .name = "Meiryo UI"
        .Size = 9
    End With

    ' 1�s�ڂ̃Z���ɍ��ږ���ݒ�
    ws.Cells(1, 1).Value = "No"
    ws.Cells(1, 2).Value = "�V�[�g��"
    ws.Cells(1, 3).Value = "���l"
    
    ' �e�[�u���̍쐬�ƃX�^�C���K�p
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C2"), , xlYes)
    ' �e�[�u���X�^�C����K�p
    lo.TableStyle = "TableStyleMedium2"
    ' �I�[�g�t�B���^�ݒ�
    lo.ShowAutoFilter = True
    ' �Ȗ͗l
    lo.ShowTableStyleRowStripes = True
    ' ���v�s
    lo.ShowTotals = False
    ' �w�b�_�s�w�i�F
    lo.HeaderRowRange.Interior.Color = RGB(64, 84, 106)
    ' �ŏ��̗�ɓ��ʂȏ�����K�p����
    lo.ShowTableStyleFirstColumn = False
    ' �Ō�̗�ɓ��ʂȏ�����K�p����
    lo.ShowTableStyleLastColumn = False

    ' �E�B���h�E�g��B2�Z���ŌŒ�
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
    
    ' �V�[�g���o��
    Dim row As Integer
    row = 2
    For Each sheet In Worksheets
        If sheet.Visible = True Then
            ws.Cells(row, 2).Value = sheet.name
            row = row + 1
        End If
    Next
    
    ' �񕝂���������
    ws.Columns("A:H").AutoFit
    ws.Columns("C").ColumnWidth = 40

End Sub

Sub ���O�̒�`�ꗗ()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = Sheets.Add

    ws.name = "���O�̒�`�ꗗ"

    ' �V�[�g�S�̂̃t�H���g��ݒ�
    With ws.Cells.Font
        .name = "Meiryo UI"
        .Size = 9
    End With

    ' 1�s�ڂ̃Z���ɍ��ږ���ݒ�
    ws.Cells(1, 1).Value = "No"
    ws.Cells(1, 2).Value = "name"
    ws.Cells(1, 3).Value = "Value"
    ws.Cells(1, 4).Value = "RefersTo"
    ws.Cells(1, 5).Value = "Visible"
    ws.Cells(1, 6).Value = "���l"
    
    ' �e�[�u���̍쐬�ƃX�^�C���K�p
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F2"), , xlYes)
    ' �e�[�u���X�^�C����K�p
    lo.TableStyle = "TableStyleMedium2"
    ' �I�[�g�t�B���^�ݒ�
    lo.ShowAutoFilter = True
    ' �Ȗ͗l
    lo.ShowTableStyleRowStripes = True
    ' ���v�s
    lo.ShowTotals = False
    ' �w�b�_�s�w�i�F
    lo.HeaderRowRange.Interior.Color = RGB(64, 84, 106)
    ' �ŏ��̗�ɓ��ʂȏ�����K�p����
    lo.ShowTableStyleFirstColumn = False
    ' �Ō�̗�ɓ��ʂȏ�����K�p����
    lo.ShowTableStyleLastColumn = False

    ' �E�B���h�E�g��B2�Z���ŌŒ�
    With Application.ActiveWindow
        .FreezePanes = False
        ws.Activate
        ws.Range("B2").Select
        .FreezePanes = True
    End With

    With ws.Columns("A:F")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.Color = RGB(192, 192, 192)
    End With

    ws.Columns("G:XFD").Hidden = True
    ws.Range("A2").Formula = "=ROW()-1"
    
    ' �V�[�g���o��
    Dim row As Integer
    row = 2
    ' ���O�̒�`�̌��������[�v
    Dim name As Object
    For Each name In Names
        ws.Cells(row, 2).Value = name.name
        ws.Cells(row, 3).Value = name.Value
        ws.Cells(row, 4).Value = name.RefersTo
        ws.Cells(row, 5).Value = name.Visible
        row = row + 1
    Next
    
    ' �񕝂���������
    ws.Columns("A:H").AutoFit
    ws.Columns("F").ColumnWidth = 40

End Sub

