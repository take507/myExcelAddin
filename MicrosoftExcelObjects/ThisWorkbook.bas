Private WithEvents m_Application As Application

' アドインインストール時処理
Private Sub Workbook_AddinInstall()
    Set m_Application = Application
    Call addinInstall
End Sub

' アドインアンインストール時処理
Private Sub Workbook_AddinUninstall()
    Call addinUninstall
End Sub

' ブックオープン時処理
Private Sub Workbook_Open()

    Set m_Application = Application
    addinInstall

    'F1キーを無効にする
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", ""

    ThisWorkbook.Saved = True
End Sub


Private Sub m_Application_WorkbookOpen(ByVal Wb As Workbook)
    ' 保護ビューで開かれている場合
    If Wb Is Nothing Or ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    Dim book As Workbook
    Set book = Workbooks(Wb.name)
    Call シートの整理(book)

    book.Saved = True
End Sub


' ブッククローズ前処理
Private Sub m_Application_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' 保護ビューで開かれている場合
    If Wb Is Nothing Or ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    ' 読み取り専用で開いている場合
    If Wb.ReadOnly = True Then
        Exit Sub
    End If

    ' 前回保存後に修正されている場合
    If Wb.Saved = False Then
        Dim book As Workbook
        Set book = Workbooks(Wb.name)
        Call シートの整理(book)
    End If

End Sub

