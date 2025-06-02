Function SVNlog()

    Dim cmd As String
    cmd = ""
    cmd = cmd & "TortoiseProc.exe"
    cmd = cmd & " /command:log"
    cmd = cmd & " /path " & ActiveWorkbook.FullName

    Call Shell(cmd, vbNormalFocus)

End Function
