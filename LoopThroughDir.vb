
Sub LoopThroughFiles()
    Dim inputDirectoryToScanForFile, filenameCriteria As String
    Dim StrFile As String
    Dim resultRow As Long
    Dim datoteka As String
    
    resultRow = ActiveCell.Row
    
    inputDirectoryToScanForFile = "C:\Edukacija\Excel\Words\"
    filenameCriteria = ".docx"

    StrFile = Dir(inputDirectoryToScanForFile & "\*" & filenameCriteria)
    Do While Len(StrFile) > 0
        
        datoteka = inputDirectoryToScanForFile & "\" & StrFile
        'Call ImportWordTable(datoteka, resultRow, StrFile)
        resultRow = resultRow + 1
        
        StrFile = Dir

    Loop

End Sub