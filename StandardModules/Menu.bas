
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
            End With
            ' �r��
            cmdbar.Controls(2).BeginGroup = True
        End If

        If cmdbar.name = "Worksheet Menu Bar" Then
            With cmdbar.Controls.Add(Type:=msoControlButton)
                .caption = "�t�@�C���̏ꏊ���J��"
                .OnAction = "�t�@�C���̏ꏊ���J��"
                .Style = msoButtonCaption
            End With
            With cmdbar.Controls.Add(Type:=msoControlButton)
                .caption = "�t�@�C���p�X���N���b�v�{�[�h�ێ�"
                .OnAction = "�t�@�C���p�X���N���b�v�{�[�h�ێ�"
                .Style = msoButtonCaption
            End With
            With cmdbar.Controls.Add(Type:=msoControlButton)
                .caption = "SVN���O"
                .OnAction = "SVN���O"
                .Style = msoButtonCaption
            End With
        End If
    Next
End Sub

' �A�h�C���̃A���C���X�g�[��
Sub addinUninstall()

    For Each cmdbar In Application.CommandBars
        If cmdbar.name = "Cell" Or cmdbar.name = "List Range Popup" Or cmdbar.name = "Worksheet Menu Bar" Then
            cmdbar.Reset
        End If
        
    Next
    
    Set cbrCmd = Nothing
End Sub