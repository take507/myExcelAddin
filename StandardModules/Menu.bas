
' アドインのインストール
Sub addinInstall()

    'On Error GoTo ErrHand
    Call addinUninstall

    For Each cmdbar In Application.CommandBars
        If cmdbar.name = "Cell" Or cmdbar.name = "List Range Popup" Then
            ' 右クリックメニューに追加
            With cmdbar.Controls.Add(Type:=msoControlPopup, before:=1)
                
                .caption = "吹き出しツール"
                With .Controls.Add
                    .caption = "吹き出し追加_赤"
                    .FaceId = 274
                    .OnAction = "吹き出し追加_赤"
                End With
                With .Controls.Add
                    .caption = "吹き出し追加_黄"
                    .FaceId = 274
                    .OnAction = "吹き出し追加_黄"
                End With
                With .Controls.Add
                    .caption = "吹き出し追加_青"
                    .FaceId = 274
                    .OnAction = "吹き出し追加_青"
                End With
                With .Controls.Add
                    .caption = "吹き出し追加_緑"
                    .FaceId = 274
                    .OnAction = "吹き出し追加_緑"
                End With
                With .Controls.Add
                    .caption = "吹き出し追加_紫"
                    .FaceId = 274
                    .OnAction = "吹き出し追加_紫"
                End With
                With .Controls.Add
                    .caption = "吹き出し一覧更新"
                    .FaceId = 274
                    .OnAction = "吹き出し一覧更新"
                    .BeginGroup = True
                End With
            End With
            With cmdbar.Controls.Add(Type:=msoControlPopup, before:=2)
                .caption = "汎用ツール"
                With .Controls.Add
                    .caption = "赤枠追加"
                    .OnAction = "赤枠追加"
                End With
                With .Controls.Add
                    .caption = "赤矢印追加"
                    .OnAction = "赤矢印追加"
                End With
                With .Controls.Add
                    .caption = "網掛け追加"
                    .OnAction = "網掛け追加"
                End With
            End With
            ' 罫線
            cmdbar.Controls(3).BeginGroup = True
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