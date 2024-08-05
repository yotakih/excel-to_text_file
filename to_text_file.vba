Sub ExportExcelToText(flderPath As String)
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim textFile As Object
    Dim fso As Object
    Dim textFilePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        textFilePath = ThisWorkbook.Path & Replace(fileName, ".xlsx", ".txt")
        Set textFile = fso.CreateTextFile(textFilePath, True)
        For Each ws In wb.Worksheets
            For Each cell In ws.UsedRange
                textFile.WriteLine cell.Address & vbTab & cell.Value
            Next cell
        Next ws
        textFile.Close
        wb.Close SaveChanges:=False
        fileName = Dir
    Loop
    
    Set textFile = Nothing
    Set fso = Nothing
End Sub

