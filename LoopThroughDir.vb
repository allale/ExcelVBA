
Sub LoopThroughFiles()
    Dim inputDirectoryToScanForFile, filenameCriteria As String
    Dim StrFile As String
    Dim resultRow As Long
    Dim datoteka As String
    
    resultRow = ActiveCell.Row
    
    'inputDirectoryToScanForFile = "C:\Edukacija\Excel\Words\"
    inputDirectoryToScanForFile = GetFolder()
    inputDirectoryToScanForFile = inputDirectoryToScanForFile & "\"
    
    'Debug.Print (inputDirectoryToScanForFile)
    
    filenameCriteria = ".docx"

    StrFile = Dir(inputDirectoryToScanForFile & "\*" & filenameCriteria)
    Do While Len(StrFile) > 0
        
        datoteka = inputDirectoryToScanForFile & "\" & StrFile
        'Call ImportWordTable(datoteka, resultRow, StrFile)
        resultRow = resultRow + 1
        
        StrFile = Dir

    Loop

End Sub

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function