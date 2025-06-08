
' アドインのインストール
Sub addinInstall()

    'On Error GoTo ErrHand
    Call addinUninstall

    For Each cmdbar In Application.CommandBars
        If cmdbar.name = "Cell" Or cmdbar.name = "List Range Popup" Then
            With cmdbar.Controls.Add(Type:=msoControlPopup, before:=1)
                .caption = "汎用ツール"
                With .Controls.Add
                    .caption = "赤枠追加"
                    .OnAction = "赤枠追加"
                End With
                With .Controls.Add
                    .caption = "赤矢印追加"
                    .OnAction = "赤矢印追加"
                End With
            End With
            ' 罫線
            cmdbar.Controls(2).BeginGroup = True
        End If

        If cmdbar.name = "Worksheet Menu Bar" Then
            With cmdbar.Controls.Add(Type:=msoControlButton)
                .caption = "ファイルの場所を開く"
                .OnAction = "ファイルの場所を開く"
                .Style = msoButtonCaption
            End With
            With cmdbar.Controls.Add(Type:=msoControlButton)
                .caption = "ファイルパスをクリップボード保持"
                .OnAction = "ファイルパスをクリップボード保持"
                .Style = msoButtonCaption
            End With
            With cmdbar.Controls.Add(Type:=msoControlButton)
                .caption = "SVNログ"
                .OnAction = "SVNログ"
                .Style = msoButtonCaption
            End With
        End If
    Next
End Sub

' アドインのアンインストール
Sub addinUninstall()

    For Each cmdbar In Application.CommandBars
        If cmdbar.name = "Cell" Or cmdbar.name = "List Range Popup" Or cmdbar.name = "Worksheet Menu Bar" Then
            cmdbar.Reset
        End If
        
    Next
    
    Set cbrCmd = Nothing
End Sub