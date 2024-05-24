Sub ImportCSV()
  
    Dim filePath As String
    Dim targetWS As Worksheet
    Dim targetRange As Range
    Dim wb As Workbook
    Dim qt As QueryTable
    Dim queryName As String
    Dim wsName As String
    wsPasteArea = "Initial Paste Area"

    filePath = "C:\Users\seedy\Downloads\orders_export.csv"
    
    Set targetWS = ThisWorkbook.Worksheets(wsPasteArea)
    targetWS.Unprotect
    Set targetRange = targetWS.Range("A1")
    
    
    With targetWS.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=targetRange)
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        
        ' UTF-8 encoding
        .TextFilePlatform = 65001
        
        .TextFileConsecutiveDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileOtherDelimiter = "|"
        .Refresh
    End With
    
    For Each qt In targetWS.QueryTables
        queryName = qt.Name
        Exit For
    Next qt
    
    On Error Resume Next
    Set QueryTable = targetWS.QueryTables(queryName)
    On Error GoTo 0
    
    If Not QueryTable Is Nothing Then
        QueryTable.Delete
    End If
    
End Sub
