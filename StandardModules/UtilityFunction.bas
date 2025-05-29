'クリッブボードへのコビー
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

' Sheets に指定した名前のシートが存在するか判定する
Public Function ExistsSheet(ByVal sheetName As String)
    Dim ws As Variant
    For Each ws In Sheets
        If LCase(ws.name) = LCase(sheetName) Then
            ExistsSheet = True '存在する
            Exit Function
        End If
    Next
    '存在しない
    ExistsSheet = False
End Function
' Sheets に指定した名前のシートが存在するか判定する
Public Function ExistsShape(sheet As Worksheet, shapeName As String)
    For Each objShp In sheet.Shapes
        If objShp.name = shapeName Then
            ExistsShape = True '存在する
            Exit Function
        End If
    Next
    '存在しない
    ExistsShape = False
End Function
' コマンドを実行
Public Function RunCmd(cmd As String) As String
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim ret As Long

    ret = wsh.Run("%ComSpec% /c " & cmd, 0, True)

    Set wsh = Nothing
End Function
