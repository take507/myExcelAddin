Private WithEvents m_Application As Application

' �A�h�C���C���X�g�[��������
Private Sub Workbook_AddinInstall()
    Set m_Application = Application
    Call addinInstall
End Sub

' �A�h�C���A���C���X�g�[��������
Private Sub Workbook_AddinUninstall()
    Call addinUninstall
End Sub

' �u�b�N�I�[�v��������
Private Sub Workbook_Open()

    Set m_Application = Application
    addinInstall

    'F1�L�[�𖳌��ɂ���
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", ""

    ThisWorkbook.Saved = True
End Sub


Private Sub m_Application_WorkbookOpen(ByVal Wb As Workbook)
    ' �ی�r���[�ŊJ����Ă���ꍇ
    If Wb Is Nothing Or ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    Dim book As Workbook
    Set book = Workbooks(Wb.name)
    Call �V�[�g�̐���(book)

    book.Saved = True
End Sub


' �u�b�N�N���[�Y�O����
Private Sub m_Application_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' �ی�r���[�ŊJ����Ă���ꍇ
    If Wb Is Nothing Or ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    ' �ǂݎ���p�ŊJ���Ă���ꍇ
    If Wb.ReadOnly = True Then
        Exit Sub
    End If

    ' �O��ۑ���ɏC������Ă���ꍇ
    If Wb.Saved = False Then
        Dim book As Workbook
        Set book = Workbooks(Wb.name)
        Call �V�[�g�̐���(book)
    End If

End Sub

