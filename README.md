Sub MergeCSVFiles()
    Dim FolderPath As String
    Dim SelectedFiles() As Variant
    Dim CurrentFile As Variant
    Dim TargetSheet As Worksheet
    Dim LastColumn As Integer
    
    'フォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVファイルが格納されているフォルダを選択してください"
        If .Show = -1 Then
            FolderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    'マージするファイル名を設定
    SelectedFiles = Array("たなか.csv", " sso_d 01 .csv", "たかし.csv", "td33333333333333333o.csv")
    
    '貼り付け先のシートを設定
    Set TargetSheet = ThisWorkbook.Sheets("貼り付け")
    
    '各ファイルを順にマージ
    For Each CurrentFile In SelectedFiles
        'ファイルが存在するか確認
        If Dir(FolderPath & "\" & CurrentFile) <> "" Then
            'CSVファイルを開く
            With Workbooks.Open(FolderPath & "\" & CurrentFile)
                'A列からC列をコピーし、新しい列に貼り付け
                LastColumn = TargetSheet.Cells(1, Columns.Count).End(xlToLeft).Column
                .Sheets(1).Range("A:C").Copy Destination:=TargetSheet.Cells(1, LastColumn + 1)
                .Close False
            End With
        End If
    Next CurrentFile
    
    '結果を表示
    MsgBox "CSVファイルのマージが完了しました。", vbInformation
End Sub



Sub レイアウトチェックマクロ()
    Dim selectedFiles As FileDialog
    Dim file As Variant
    Dim workFile As Workbook
    Dim pasteSheet As Worksheet
    Dim copyRange1 As Range
    Dim copyRange2 As Range
    Dim saveFileName As String
    
    ' ①エクセルファイルを選択
    Set selectedFiles = Application.FileDialog(msoFileDialogOpen)
    selectedFiles.AllowMultiSelect = True
    selectedFiles.Title = "エクセルファイルを選択してください"
    
    If selectedFiles.Show = -1 Then
        ' 選択されたファイル数分繰り返し
        For Each file In selectedFiles.SelectedItems
            ' ②選択されたエクセルファイルを開く
            Set workFile = Workbooks.Open(file)
            
            ' 作業シートを取得（貼り付けシート）
            Set pasteSheet = ThisWorkbook.Sheets("貼り付けシート")
            
            ' 貼り付け先セルの範囲を指定
            Set copyRange1 = workFile.Sheets(1).Range("C3:C5")
            Set copyRange2 = workFile.Sheets(1).Range("C10:C15")
            
            ' 貼り付け
            pasteSheet.Range("B1").Value = copyRange1.Value
            pasteSheet.Range("D1").Value = copyRange2.Value
            
            ' 作業ファイルを閉じる
            workFile.Close SaveChanges:=False
            
            ' ③ファイルを保存
            saveFileName = "レイアウトチェック_" & pasteSheet.Range("B1").Value
            ThisWorkbook.SaveCopyAs "D:\" & saveFileName
            
            ' 作業シートをクリア
            pasteSheet.Range("B1:D1").ClearContents
        Next file
    End If
    
    ' ④繰り返し実行終了時の処理
    MsgBox "処理が完了しました。"
End Sub

Sub レイアウトチェックマクロ()

    Dim selectedFiles As FileDialog
    Dim file As Variant
    Dim workFile As Workbook
    Dim pasteSheet As Worksheet
    Dim copyRange1 As Range
    Dim copyRange2 As Range
    Dim saveFileName As String
    
    ' ①エクセルファイルを選択
    Set selectedFiles = Application.FileDialog(msoFileDialogOpen)
    selectedFiles.AllowMultiSelect = True
    selectedFiles.Title = "エクセルファイルを選択してください"
    
    If selectedFiles.Show = -1 Then
        ' 選択されたファイル数分繰り返し
        For Each file In selectedFiles.SelectedItems
            ' ②選択されたエクセルファイルを開く
            Set workFile = Workbooks.Open(file)
            
            ' 作業シートを取得（貼り付けシート）
            Set pasteSheet = ThisWorkbook.Sheets("貼り付けシート")
            
            ' 貼り付け先セルの範囲を指定
            Set copyRange1 = workFile.Sheets(1).Range("C3:C5")
            Set copyRange2 = workFile.Sheets(1).Range("C10:C15")
            
            ' 貼り付け
            pasteSheet.Range("B1").Value = copyRange1.Value
            pasteSheet.Range("D1").Value = copyRange2.Value
            
            ' 作業ファイルを閉じる
            workFile.Close SaveChanges:=False
            
            ' ③ファイルをマクロ有効なエクセルファイルとして保存
            saveFileName = "レイアウトチェック_" & pasteSheet.Range("B1").Value & ".xlsm"
            ThisWorkbook.SaveCopyAs "D:\" & saveFileName
            
            ' 作業シートをクリア
            pasteSheet.Range("M5:M28").ClearContents
            pasteSheet.Range("M31:M68").ClearContents
        Next file
    End If
    
    ' ④繰り返し実行終了時の処理
    MsgBox "処理が完了しました。"
End Sub



