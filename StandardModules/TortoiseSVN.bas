Function SVN���O()

    Dim cmd As String
    cmd = ""
    cmd = cmd & "TortoiseProc.exe"
    cmd = cmd & " /command:log"
    cmd = cmd & " /path " & ActiveWorkbook.Path

    Call Shell(cmd, vbNormalFocus)
    

End Function
