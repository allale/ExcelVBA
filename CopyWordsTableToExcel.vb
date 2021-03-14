Option Explicit


Sub LoopThroughFiles()
    Dim inputDirectoryToScanForFile, filenameCriteria As String
    Dim StrFile As String
    Dim resultRow As Long
    Dim datoteka As String
    
    resultRow = ActiveCell.Row
    
    inputDirectoryToScanForFile = "C:\Edukacija\Excel\Words\"
    'filenameCriteria = ".docx"

    StrFile = Dir(inputDirectoryToScanForFile & "\*" & filenameCriteria)
    Do While Len(StrFile) > 0
        
        datoteka = inputDirectoryToScanForFile & "\" & StrFile
        Call ImportWordTable(datoteka, resultRow, StrFile)
        resultRow = resultRow + 1
        
        StrFile = Dir

    Loop

End Sub


Sub ImportWordTable(wdFileName As Variant, resultRow As Long, imeKolegija As String)

Dim wdDoc As Object
Dim tableNo As Integer      'table number in Word
Dim iRow As Long            'row index in Excel
Dim iCol As Integer         'column index in Excel
Dim tableStart As Integer
Dim tableTot As Integer


On Error Resume Next


Set wdDoc = GetObject(wdFileName) 'open Word file

With wdDoc
    tableNo = wdDoc.tables.Count
    tableTot = wdDoc.tables.Count
    If tableNo = 0 Then
        MsgBox "This document contains no tables", _
        vbExclamation, "Import Word Table"
    ElseIf tableNo > 1 Then
        tableNo = 4
    End If

    'resultRow = ActiveCell.Row
    Cells(resultRow, 1).Value = imeKolegija
    resultRow = resultRow + 1

    For tableStart = tableNo To tableNo
        With .tables(tableStart)
            'copy cell contents from Word table cells to Excel cells
            For iRow = 1 To .Rows.Count
                For iCol = 1 To .Columns.Count
                    Cells(resultRow, iCol).NumberFormat = "@"
                    Cells(resultRow, iCol) = WorksheetFunction.Clean(.cell(iRow, iCol).Range.text)
                Next iCol
                resultRow = resultRow + 1
            Next iRow
        End With
        resultRow = resultRow + 1
    Next tableStart
End With

End Sub

