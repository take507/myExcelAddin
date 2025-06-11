Function 吹き出し作成() As shape

    Dim shape As shape
    Set shape = ActiveSheet.Shapes.AddShape(msoShapeLineCallout1, _
        ActiveCell.Left, ActiveCell.Top, _
        200, 60)

    With shape
        '背景色
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        '罫線色
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        '文字色
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Adjustments.Item(1) = 0
        .Adjustments.Item(2) = 0
        .Adjustments.Item(3) = -0.2856
        .Adjustments.Item(4) = -0.15919

        '入力文字
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & "XXXX"

        'フォント
        With .TextFrame2.TextRange.Font
            .Size = 9
            .NameComplexScript = "メイリオ"
            .NameFarEast = "メイリオ"
            .name = "メイリオ"
        End With

        'テキストマージン
        With .TextFrame2
            .MarginTop = 2.8346456693
            .MarginLeft = 7.0866141732
            .MarginRight = 2.8346456693
            .MarginBottom = 2.8346456693
        End With

        'オブジェクト名
        .name = "revShape_" & Format(Now(), "YYYY-MM-DD_HH:MM:SS") & "_" & Environ("USERNAME") & "_" & shape.ID
    End With

    Set 吹き出し作成 = shape
End Function

Sub 吹き出し追加_赤()

    Dim shape As shape
    Set shape = 吹き出し作成

    With shape
        '背景色
        .Fill.ForeColor.RGB = RGB(255, 204, 204)
        '罫線色
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        '文字色
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
        '入力文字
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub 吹き出し追加_黄()

    Dim shape As shape
    Set shape = 吹き出し作成

    With shape
        '背景色
        .Fill.ForeColor.RGB = RGB(255, 255, 204)
        '罫線色
        .Line.ForeColor.RGB = RGB(191, 114, 0)
        '文字色
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        '入力文字
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub 吹き出し追加_青()

    Dim shape As shape
    Set shape = 吹き出し作成

    With shape
        '背景色
        .Fill.ForeColor.RGB = RGB(218, 227, 243)
        '罫線色
        .Line.ForeColor.RGB = RGB(0, 176, 240)
        '文字色
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 112, 192)
        '入力文字
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub 吹き出し追加_緑()

    Dim shape As shape
    Set shape = 吹き出し作成

    With shape
        '背景色
        .Fill.ForeColor.RGB = RGB(226, 240, 217)
        '罫線色
        .Line.ForeColor.RGB = RGB(84, 130, 53)
        '文字色
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(84, 130, 53)
        '入力文字
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub
Sub 吹き出し追加_紫()

    Dim shape As shape
    Set shape = 吹き出し作成

    With shape
        '背景色
        .Fill.ForeColor.RGB = RGB(220, 185, 255)
        '罫線色
        .Line.ForeColor.RGB = RGB(112, 48, 169)
        '文字色
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(133, 61, 13)
        '入力文字
        .TextFrame2.TextRange.Characters.text = "" & Chr(13) & ""
    End With

End Sub

Sub CreateNewSheetWithTableAndAllSettingsFiltered()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = Sheets.Add

    ws.name = "吹き出し一覧"

    ' シート全体のフォントを設定
    With ws.Cells.Font
        .name = "Meiryo UI"
        .Size = 9
    End With

    ' 1行目のセルに項目名を設定
    ws.Cells(1, 1).Value = "No"
    ws.Cells(1, 2).Value = "ID"
    ws.Cells(1, 3).Value = "内容"
    ws.Cells(1, 4).Value = "対象シート"
    ws.Cells(1, 5).Value = "対応内容"
    ws.Cells(1, 6).Value = "対応区分"
    ws.Cells(1, 7).Value = "確認日"
    ws.Cells(1, 8).Value = "確認者"

    ' 列幅を自動調整
    ws.Columns("A:H").AutoFit

    ' テーブルの作成とスタイル適用
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H1"), , xlYes)

    ' テーブルスタイルを適用
    lo.TableStyle = "TableStyleMedium2"
    ' オートフィルタ設定
    lo.ShowAutoFilter = True
    ' 縞模様
    lo.ShowTableStyleRowStripes = True
    ' 合計行
    lo.ShowTotals = False
    ' 最初の列に特別な書式を適用する
    lo.ShowTableStyleFirstColumn = True
    ' ヘッダ行背景色
    lo.HeaderRowRange.Interior.Color = RGB(64, 84, 106)
    ' 最後の列に特別な書式を適用する
    lo.ShowTableStyleLastColumn = False

    ' ウィンドウ枠をB2セルで固定
    With Application.ActiveWindow
        .FreezePanes = False
        ws.Activate
        ws.Range("B2").Select
        .FreezePanes = True
    End With

    With ws.Columns("A:H")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.Color = RGB(192, 192, 192)
    End With

    ws.Columns("I:XFD").Hidden = True
    ws.Range("A2").Formula = "=ROW()-1"

End Sub

Sub 吹き出し一覧更新()

    Application.ScreenUpdating = False

    Const SHEET_NAME_REV As String = "吹き出し一覧"
    Const COL_SHAPE_ID As Integer = 2
    Const COL_REVIEW_COMMENT As Integer = 3
    Const COL_SHEET_NAME As Integer = 4
    Const COL_DETAIL As Integer = 5
    Const COL_KIND As Integer = 6
    Const COL_CHECK_DATE As Integer = 7
    Const COL_CHECK_ACTOR As Integer = 8
    Const ROW_FIRST As Integer = 3


    If ExistsSheet(SHEET_NAME_REV) = False Then
        Call CreateNewSheetWithTableAndAllSettingsFiltered
    End If

    Dim revSheet As Worksheet
    Set revSheet = Worksheets(SHEET_NAME_REV)

    'リストオブジェクトが存在しない場合
    If revSheet.ListObjects.Count = 0 Then Exit Sub

    Set listObj = revSheet.ListObjects(1)

    Dim i As Long
    Dim shape As shape
    Dim addCnt As Long
    Dim delCnt As Long
    addCnt = 0
    delCnt = 0

    'シートの件数分ループ
    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Worksheets
        'オートシェイプの件数分ループ
        For Each shape In sheet.Shapes

            '吹き出し追加で追加されたオブジェクトの場合
            If shape.name Like "revShape_*" Then

                Dim findF As Boolean
                findF = False

                'オブジェクト名の中のshapeIDが不一致の場合（吹き出しをコピーした場合）
                If Not (shape.name Like "revShape_*_" & shape.ID) Then
                    Debug.Print shape.name & " " & shape.ID
                    Dim ary As Variant
                    ary = Split(shape.name, "_")
                    Dim oldId As String
                    oldId = ary(UBound(ary))
                    shape.name = "revShape_" & Format(Now(), "YYYY-MM-DD_HH:MM:SS") & "_copied" & oldId & "_" & shape.ID
                End If

                'リストの件数分繰り返し処理
                For i = 1 To listObj.ListRows.Count
                    'オブジェクト名が一致した場合
                    If listObj.ListRows(i).Range(COL_SHAPE_ID).Value = shape.name Then
                        'フラグを立てる
                        findF = True
                        Exit For
                    End If
                Next i

                '同名キーが見つからなかった場合
                If findF = False Then
                    'レコード追加し登録する
                    With listObj.ListRows.Add
                        .Range(COL_SHAPE_ID).Value = shape.name
                        .Range(COL_REVIEW_COMMENT).Value = shape.TextFrame2.TextRange.Characters.text
                        .Range(COL_SHEET_NAME).Value = "=HYPERLINK(""#" + sheet.name + "!" + shape.TopLeftCell.Address + """,""" + sheet.name + """)"
                    End With
                    addCnt = addCnt + 1
                End If


                '空行を削除
                For i = listObj.ListRows.Count To 1 Step -1
                    If listObj.ListRows(i).Range(COL_SHAPE_ID) = "" Then
                        listObj.ListRows(i).Delete
                    End If
                Next i

            End If
        Next
    Next

    'リストの件数分繰り返し処理
    For i = 1 To listObj.ListRows.Count

        With listObj.ListRows(i)
            '確認日が記載されている and 確認者が記載されている
            '対応区分が「対応済み」「対応不要」「重複」の場合
            If Len(.Range(COL_CHECK_DATE).Value) > 0 And Len(.Range(COL_CHECK_ACTOR).Value) > 0 And _
               (.Range(COL_KIND).Value = "対応済み" Or .Range(COL_KIND).Value = "対応不要" Or .Range(COL_KIND).Value = "重複") Then

                'シート名取得
                Dim sheetName As String
                sheetName = listObj.ListRows(i).Range(COL_SHEET_NAME).Value
                'shape名取得
                Dim shapeName As String
                shapeName = listObj.ListRows(i).Range(COL_SHAPE_ID).Value

                'シートもshapeも存在する場合
                If ExistsSheet(sheetName) = True Then
                    Set sheet = Worksheets(listObj.ListRows(i).Range(COL_SHEET_NAME).Value)
                    If ExistsShape(sheet, shapeName) = True Then
                        'shapeを削除する
                        sheet.Shapes(shapeName).Delete
                        delCnt = delCnt + 1
                    End If
                End If

            End If
        End With

    Next

    Worksheets(1).Select
    Application.ScreenUpdating = True
    Worksheets(SHEET_NAME_REV).Select

    Dim msg As String
    msg = ""
    If addCnt > 0 Then
        msg = "メモ一覧に" & addCnt & "件追加しました。" & Chr(10)
    End If
    If delCnt > 0 Then
        msg = msg & "吹き出しを" & delCnt & "件削除しました。"
    End If
    If Len(msg) > 0 Then
        MsgBox (msg)
    End If

End Sub
