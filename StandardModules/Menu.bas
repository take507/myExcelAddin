
' �A�h�C���̃C���X�g�[��
Sub addinInstall()

    'On Error GoTo ErrHand
    Call addinUninstall

    For Each cmdbar In Application.CommandBars
        If cmdbar.name = "Cell" Or cmdbar.name = "List Range Popup" Then
            With cmdbar.Controls.Add(Type:=msoControlPopup, before:=1)
                .caption = "�ėp�c�[��"
                With .Controls.Add
                    .caption = "�Ԙg�ǉ�"
                    .OnAction = "�Ԙg�ǉ�"
                End With
                With .Controls.Add
                    .caption = "�Ԗ��ǉ�"
                    .OnAction = "�Ԗ��ǉ�"
                End With
                With .Controls.Add
                    .caption = "�t�@�C���̏ꏊ���J��"
                    .OnAction = "�t�@�C���̏ꏊ���J��"
                End With
                With .Controls.Add
                    .caption = "�t�@�C���p�X���N���b�v�{�[�h�ێ�"
                    .OnAction = "�t�@�C���p�X���N���b�v�{�[�h�ێ�"
                End With
            End With
            ' �r��
            cmdbar.Controls(2).BeginGroup = True
        End If
    Next
End Sub

' �A�h�C���̃A���C���X�g�[��
Sub addinUninstall()

    For Each cmdbar In Application.CommandBars
        If cmdbar.name = "Cell" Or cmdbar.name = "List Range Popup" Then
            cmdbar.Reset
        End If
    Next

    Set cbrCmd = Nothing
End Sub
