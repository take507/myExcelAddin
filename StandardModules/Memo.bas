Function �����o���쐬() As shape

    Dim shape As shape
    Set shape = ActiveSheet.Shapes.AddShape(msoShapeLineCallout1, _
        ActiveCell.Left, ActiveCell.Top, _
        200, 60)

    With shape
        '�w�i�F
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        '�r���F
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        '�����F
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Adjustments.Item(1) = 0
        .Adjustments.Item(2) = 0
        .Adjustments.Item(3) = -0.2856
        .Adjustments.Item(4) = -0.15919

        '���͕���
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & "XXXX"

        '�t�H���g
        With .TextFrame2.TextRange.Font
            .Size = 9
            .NameComplexScript = "���C���I"
            .NameFarEast = "���C���I"
            .name = "���C���I"
        End With

        '�e�L�X�g�}�[�W��
        With .TextFrame2
            .MarginTop = 2.8346456693
            .MarginLeft = 7.0866141732
            .MarginRight = 2.8346456693
            .MarginBottom = 2.8346456693
        End With

        '�I�u�W�F�N�g��
        .name = "revShape_" & Format(Now(), "YYYY-MM-DD_HH:MM:SS") & "_" & Environ("USERNAME") & "_" & shape.ID
    End With

    Set �����o���쐬 = shape
End Function

Sub �����o���ǉ�_1()

    Dim shape As shape
    Set shape = �����o���쐬

    With shape
        '�w�i�F
        .Fill.ForeColor.RGB = RGB(255, 204, 204)
        '�r���F
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        '�����F
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
        '���͕���
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub �����o���ǉ�_2()

    Dim shape As shape
    Set shape = �����o���쐬

    With shape
        '�w�i�F
        .Fill.ForeColor.RGB = RGB(255, 255, 204)
        '�r���F
        .Line.ForeColor.RGB = RGB(191, 114, 0)
        '�����F
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        '���͕���
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub �����o���ǉ�_3()

    Dim shape As shape
    Set shape = �����o���쐬

    With shape
        '�w�i�F
        .Fill.ForeColor.RGB = RGB(218, 227, 243)
        '�r���F
        .Line.ForeColor.RGB = RGB(0, 176, 240)
        '�����F
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 112, 192)
        '���͕���
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub �����o���ǉ�_4()

    Dim shape As shape
    Set shape = �����o���쐬

    With shape
        '�w�i�F
        .Fill.ForeColor.RGB = RGB(226, 240, 217)
        '�r���F
        .Line.ForeColor.RGB = RGB(84, 130, 53)
        '�����F
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(84, 130, 53)
        '���͕���
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub �����o���ǉ�_5()

    Dim shape As shape
    Set shape = �����o���쐬

    With shape
        '�w�i�F
        .Fill.ForeColor.RGB = RGB(255, 204, 153)
        '�r���F
        .Line.ForeColor.RGB = RGB(237, 126, 51)
        '�����F
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(133, 61, 13)
        '���͕���
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub

Sub CreateNewSheetWithTableAndAllSettingsFiltered()
    Dim ws As Worksheet
    Dim lo As ListObject ' ListObject�̓e�[�u����\���I�u�W�F�N�g�ł�

    ' �V�����V�[�g���쐬���A�ϐ��Ɋi�[
    Set ws = Sheets.Add

    ' �V�[�g����t����ꍇ�͈ȉ��̍s�̃R�����g���O���Ă�������
    ' ws.Name = "�^�X�N�Ǘ��\"

    ' �V�[�g�S�̂̃t�H���g��ݒ�
    With ws.Cells.Font
        .name = "Meiryo UI" ' �܂��� "���C���I" �ł���
        .Size = 9
    End With

    ' 1�s�ڂ̃Z���ɍ��ږ���ݒ�
    ws.Cells(1, 1).Value = "No"
    ws.Cells(1, 2).Value = "ID"
    ws.Cells(1, 3).Value = "���e"
    ws.Cells(1, 4).Value = "�ΏۃV�[�g"
    ws.Cells(1, 5).Value = "�Ή����e"
    ws.Cells(1, 6).Value = "�Ή��敪"
    ws.Cells(1, 7).Value = "�m�F��"
    ws.Cells(1, 8).Value = "�m�F��"

    ' 1�s�ڂ̃Z���͈́iA1����H1�j�̏����ݒ�
    With ws.Range("A1:H1")
        .Font.Color = RGB(255, 255, 255) ' �����F�𔒂ɐݒ� (R=255, G=255, B=255)
        .Interior.Color = RGB(0, 0, 128) ' �w�i�F�����F�ɐݒ� (R=0, G=0, B=128)
        .Font.Bold = True ' �����𑾎��ɂ���
        .VerticalAlignment = xlVAlignCenter ' ����������������
        .HorizontalAlignment = xlHAlignCenter ' ����������������
    End With

    ' �񕝂��������� (�I�v�V����)
    ' ws.Columns("A:H").AutoFit

    ' �e�[�u���̍쐬�ƃX�^�C���K�p
    ' ���ږ����܂ޔ͈�A1:H1���e�[�u���Ƃ��Ē�`
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H1"), , xlYes)

    ' �e�[�u���ɖ��O��t���� (�C��)
    lo.name = "�^�X�N�Ǘ��e�[�u��"

    ' �e�[�u���X�^�C����K�p
    lo.TableStyle = "TableStyleMedium2" ' ��: "TableStyleMedium2"

    ' �w�b�_�[�s�̃t�B���^�[�{�^����**�\��**�ɂ��� (True)
    lo.ShowAutoFilter = True

    ' �Ȗ͗l��L���ɂ���ꍇ�� True �ɐݒ�
    lo.ShowTableStyleRowStripes = True
    ' �e�[�u���ɍ��v�s��\������ꍇ�� True �ɐݒ�
    lo.ShowTotals = False
    ' �ŏ��̗�ɓ��ʂȏ�����K�p����ꍇ�� True �ɐݒ�
    lo.ShowTableStyleFirstColumn = False
    ' �Ō�̗�ɓ��ʂȏ�����K�p����ꍇ�� True �ɐݒ�
    lo.ShowTableStyleLastColumn = False


    ' �E�B���h�E�g��B2�Z���ŌŒ�
    With Application.ActiveWindow
        .FreezePanes = False ' �����̃E�B���h�E�g�Œ������ (�O�̂���)
        ws.Activate ' �V�[�g���A�N�e�B�u�ɂ���
        ws.Range("B2").Select ' B2�Z����I��
        .FreezePanes = True ' �E�B���h�E�g���Œ�
    End With

    ' A�񂩂�H��̂��ׂẴZ���Ɍr��������
    With ws.Columns("A:H")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous ' ���r��
        .Borders(xlEdgeTop).LineStyle = xlContinuous ' ��r��
        .Borders(xlEdgeRight).LineStyle = xlContinuous ' �E�r��
        .Borders(xlEdgeBottom).LineStyle = xlContinuous ' ���r��
        .Borders(xlInsideVertical).LineStyle = xlContinuous ' ���������̓����r��
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous ' ���������̓����r��
        .Borders.Color = RGB(192, 192, 192) ' �r���̐F���D�F�ɂ��� (�C��)
    End With

    ' I��ȍ~���\���ɂ���
    ws.Columns("I:XFD").Hidden = True

    MsgBox "�V�����V�[�g���쐬����A���ׂĂ̐ݒ肪�������܂����B", vbInformation

End Sub

Sub �����ꗗ�X�V()

    Application.ScreenUpdating = False

    Const SHEET_NAME_REV As String = "�����ꗗ"
    Const COL_SHAPE_ID As Integer = 2
    Const COL_REVIEW_COMMENT As Integer = 3
    Const COL_SHEET_NAME As Integer = 4
    Const COL_DETAIL As Integer = 5
    Const COL_KIND As Integer = 6
    Const COL_CHECK_DATE As Integer = 7
    Const COL_CHECK_ACTOR As Integer = 8
    Const ROW_FIRST As Integer = 3


    If ExistsSheet(SHEET_NAME_REV) = False Then
        ThisWorkbook.Sheets(SHEET_NAME_REV).Copy After:=Worksheets(Worksheets.Count)
    End If

    Dim revSheet As Worksheet
    Set revSheet = Worksheets(SHEET_NAME_REV)

    '���X�g�I�u�W�F�N�g�����݂��Ȃ��ꍇ
    If revSheet.ListObjects.Count = 0 Then Exit Sub

    Set listObj = revSheet.ListObjects(1)

    Dim i As Long
    Dim shape As shape
    Dim addCnt As Long
    Dim delCnt As Long
    addCnt = 0
    delCnt = 0

    '�V�[�g�̌��������[�v
    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Worksheets
        '�I�[�g�V�F�C�v�̌��������[�v
        For Each shape In sheet.Shapes

            '�����o���ǉ��Œǉ����ꂽ�I�u�W�F�N�g�̏ꍇ
            If shape.name Like "revShape_*" Then

                Dim findF As Boolean
                findF = False

                '�I�u�W�F�N�g���̒���shapeID���s��v�̏ꍇ�i�����o�����R�s�[�����ꍇ�j
                If Not (shape.name Like "revShape_*_" & shape.ID) Then
                    Debug.Print shape.name & " " & shape.ID
                    Dim ary As Variant
                    ary = Split(shape.name, "_")
                    Dim oldId As String
                    oldId = ary(UBound(ary))
                    shape.name = "revShape_" & Format(Now(), "YYYY-MM-DD_HH:MM:SS") & "_copied" & oldId & "_" & shape.ID
                End If

                '���X�g�̌������J��Ԃ�����
                For i = 1 To listObj.ListRows.Count
                    '�I�u�W�F�N�g������v�����ꍇ
                    If listObj.ListRows(i).Range(COL_SHAPE_ID).Value = shape.name Then
                        '�t���O�𗧂Ă�
                        findF = True
                        Exit For
                    End If
                Next i

                '�����L�[��������Ȃ������ꍇ
                If findF = False Then
                    '���R�[�h�ǉ����o�^����
                    With listObj.ListRows.Add
                        .Range(COL_SHAPE_ID).Value = shape.name
                        .Range(COL_REVIEW_COMMENT).Value = shape.TextFrame2.TextRange.Characters.text
                        .Range(COL_SHEET_NAME).Value = "=HYPERLINK(""#" + sheet.name + "!" + shape.TopLeftCell.Address + """,""" + sheet.name + """)"
                    End With
                    addCnt = addCnt + 1
                End If


                '��s���폜
                For i = listObj.ListRows.Count To 1 Step -1
                    If listObj.ListRows(i).Range(COL_SHAPE_ID) = "" Then
                        listObj.ListRows(i).Delete
                    End If
                Next i

            End If
        Next
    Next

    '���X�g�̌������J��Ԃ�����
    For i = 1 To listObj.ListRows.Count

        With listObj.ListRows(i)
            '�m�F�����L�ڂ���Ă��� and �m�F�҂��L�ڂ���Ă���
            '�Ή��敪���u�Ή��ς݁v�u�Ή��s�v�v�u�d���v�̏ꍇ
            If Len(.Range(COL_CHECK_DATE).Value) > 0 And Len(.Range(COL_CHECK_ACTOR).Value) > 0 And _
               (.Range(COL_KIND).Value = "�Ή��ς�" Or .Range(COL_KIND).Value = "�Ή��s�v" Or .Range(COL_KIND).Value = "�d��") Then

                '�V�[�g���擾
                Dim sheetName As String
                sheetName = listObj.ListRows(i).Range(COL_SHEET_NAME).Value
                'shape���擾
                Dim shapeName As String
                shapeName = listObj.ListRows(i).Range(COL_SHAPE_ID).Value

                '�V�[�g��shape�����݂���ꍇ
                If ExistsSheet(sheetName) = True Then
                    Set sheet = Worksheets(listObj.ListRows(i).Range(COL_SHEET_NAME).Value)
                    If ExistsShape(sheet, shapeName) = True Then
                        'shape���폜����
                        sheet.Shapes(shapeName).Delete
                        delCnt = delCnt + 1
                    End If
                End If

            End If
        End With

    Next

    Worksheets(1).Select
    Application.ScreenUpdating = True
    Worksheets(SHEET_NAME_REV).Select

    Dim msg As String
    msg = ""
    If addCnt > 0 Then
        msg = "�����ꗗ��" & addCnt & "���ǉ����܂����B" & Chr(10)
    End If
    If delCnt > 0 Then
        msg = msg & "�����o����" & delCnt & "���폜���܂����B"
    End If
    If Len(msg) > 0 Then
        MsgBox (msg)
    End If

End Sub
