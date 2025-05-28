' アドインインストール時処理
Private Sub Workbook_AddinInstall()
    Set m_Application = Application
    Call addinInstall
End Sub

' アドインアンインストール時処理
Private Sub Workbook_AddinUninstall()
    Call addinUninstall
End Sub
