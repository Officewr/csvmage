Sub MergeCSVFiles()
    
    Dim myPath As String
    Dim myFile As String
    Dim myExtension As String
    Dim myWorkbook As Workbook
    Dim DestWB As Workbook
    Dim DestWS As Worksheet
    Dim LastCol As Long
    Dim FileCounter As Long
    Dim i As Long
    
    'CSVファイルが保存されたフォルダーのパスを指定します
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVファイルの保存されているフォルダを選択してください"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        myPath = .SelectedItems(1) & "\"
    End With
    
    'ファイルをマージするExcelファイルのシート名を指定します
    Set DestWB = Workbooks("集計くん.xlsx")
    Set DestWS = DestWB.Sheets("貼り付け")
    
    'CSVファイルの拡張子を指定します
    myExtension = "*.csv*"
    
    'CSVファイルのファイル名の配列を作成します
    Dim Files() As String
    Files = Array("山田.csv", "田中.csv", "木村.csv", "西田.csv")
    
    '各ファイルの行数を格納する配列を作成します
    Dim RowCounts() As Long
    RowCounts = Array(130, 130, 611, 611)
    
    '各ファイルを順番に取り込みます
    For i = 0 To UBound(Files)
        myFile = myPath & Files(i)
        Set myWorkbook = Workbooks.Open(myFile)
        If myWorkbook.Sheets(1).UsedRange.Rows.Count <> RowCounts(i) Then
            myWorkbook.Sheets(1).Cells(1, 1).Interior.Color = vbRed '指定行数と異なる場合、1番上のセルに赤色を付ける
        End If
        LastCol = DestWS.Cells(1, DestWS.Columns.Count).End(xlToLeft).Column
        For j = 1 To 3 '列数を3に変更
            FileCounter = FileCounter + 1
            myWorkbook.Sheets(1).Cells(1, j).Copy
            DestWS.Cells(1, LastCol + FileCounter).PasteSpecial xlPasteValues
        Next j
        myWorkbook.Close
    Next i
    
    'アプリケーションのコピー＆ペーストバッファをクリアします
    Application.CutCopyMode = False
    
    MsgBox "CSVファイルのマージが完了しました。"
    
End Sub
