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
