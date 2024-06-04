Sub GetDataFromWorkbooks()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim startPasteRow As Long
    Dim dataRange As Range
    
    ' 出力先ワークシートを設定
    Set destWs = ThisWorkbook.Sheets(1)
    
    ' フォルダのパスを指定
    folderPath = ThisWorkbook.Path & "\data\"
   
    ' 出力先ファイル名とファイルパスを指定
    fileResultName = ThisWorkbook.Path & "\result.xlsm" ' 新しいファイル名を設定

    ' フォルダ内の最初のExcelファイルの名前を取得
    fileName = Dir(folderPath & "*.xls*")
    
    ' フォルダ内のすべてのExcelファイルをループ処理
    Do While fileName <> ""
        ' Excelファイルを開く
        Set wb = Workbooks.Open(folderPath & fileName)
        
        ' ファイル内のワークシートを設定
        Set ws = wb.Sheets(1)
        
        ' 最終行を取得
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        
        ' データが存在する場合にのみ処理を実行
        If lastRow >= 8 Then ' 8行目からデータが開始するため
            ' 出力先の次の空白行を取得（先頭5行を空ける）
            startPasteRow = destWs.Cells(destWs.Rows.Count, 2).End(xlUp).Row + 1
            If startPasteRow < 6 Then startPasteRow = 6 ' 最低でも6行目からデータを開始する
            
            ' データ範囲を設定（No列を除く）
            Set dataRange = ws.Range("B8:M" & lastRow)
            
            ' データをコピーして貼り付ける
            dataRange.Copy
            destWs.Cells(startPasteRow, 2).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End If
        
        ' ファイルを閉じる
        wb.Close SaveChanges:=False
        
        ' フォルダ内の次のExcelファイルの名前を取得
        fileName = Dir
    Loop
    
    ' ソートを行うための最終行を取得
    lastRow = destWs.Cells(destWs.Rows.Count, 2).End(xlUp).Row

    ' B列（日付列）を基準にデータを昇順でソートする
    With destWs.Sort
        .SortFields.Clear
        .SortFields.Add Key:=destWs.Range("B7:B" & lastRow), Order:=xlAscending
        .SetRange destWs.Range("A7:M" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' 通し番号を追加
    For i = 8 To lastRow
        destWs.Cells(i, "A").Value = i - 7 ' 7行目から通し番号を振り始める
    Next i
    
    Call FormatTimeColumn()

    ' ファイルを別名で保存
    ThisWorkbook.SaveCopyAs Filename:= fileResultName

    ' データをクリアする範囲を特定する
    ' ここではB列のデータを基準にして最終行を見つけています
    lastRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    
    ' 出力データの範囲をクリアする
    destWs.Range("B8:M" & lastRow).ClearContents


    ' Unicodeの波ダッシュを使用して、D列の指定された範囲に挿入
    ' 全角波ダッシュのUnicodeは"U+301C"ですが、
    ' Windowsでは一般に"U+FF5E"を使用します。
    Dim waveDash As String
    waveDash = ChrW(&HFF5E)

    ' D列に"~"を入力する
    destWs.Range("D8:D" & lastRow).Value = waveDash
End Sub

Sub FormatTimeColumn()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' 貼り付け先ワークシートを設定
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row ' 時間データの最終行を取得

    ' 時間データが含まれるC列のセルの書式を統一する
    With ws.Range("G8:G" & lastRow)
        .NumberFormat = "h:mm;@" ' 時間の書式を "時:分" に設定
        .HorizontalAlignment = xlCenter ' セルの内容を中央寄せにする
    End With
End Sub
