'�N���b�u�{�[�h�ւ̃R�r�[
Public Function ClipBoardSave(temp As String)
'https://vba-create.jp/vba-error-clipboard-copy/
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .text = temp
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Function

' Sheets �Ɏw�肵�����O�̃V�[�g�����݂��邩���肷��
Public Function ExistsSheet(ByVal sheetName As String)
    Dim ws As Variant
    For Each ws In Sheets
        If LCase(ws.name) = LCase(sheetName) Then
            ExistsSheet = True '���݂���
            Exit Function
        End If
    Next
    '���݂��Ȃ�
    ExistsSheet = False
End Function
' Sheets �Ɏw�肵�����O�̃V�[�g�����݂��邩���肷��
Public Function ExistsShape(sheet As Worksheet, shapeName As String)
    For Each objShp In sheet.Shapes
        If objShp.name = shapeName Then
            ExistsShape = True '���݂���
            Exit Function
        End If
    Next
    '���݂��Ȃ�
    ExistsShape = False
End Function
' �R�}���h�����s
Public Function RunCmd(cmd As String) As String
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim ret As Long

    ret = wsh.Run("%ComSpec% /c " & cmd, 0, True)

    Set wsh = Nothing
End Function
