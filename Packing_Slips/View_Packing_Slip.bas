Sub PrintPackingSlips()

    Application.ScreenUpdating = False
    
    Sheets("Daily Data").Unprotect
    Sheets("Daily Data").visible = True
    Sheets("Daily Data").Range("A2:X100000").ClearContents
    
    Sheets("Packing Data").Unprotect
    Sheets("Packing Data").Range("B2:W1000").ClearContents
    
    Sheets("Shopify All Data").visible = True
    Sheets("Shopify All Data").Unprotect
    Sheets("Shopify All Data").AutoFilterMode = False
    
    Sheets("Seed Data").visible = True
    
    Sheets("Intermediate").Unprotect
    Sheets("Intermediate").Range("A2:W1000").ClearContents
    
    Dim singleOrder As Boolean
    
    'For printing a single packing slip that has previously been printed
    If Range("ONEOFF").value <> "" Then
        Sheets("Packing Slips").Range("Q7").value = Range("ONEOFF").value
        If FilterForOldOrder() = True Then
            'copy the visible cells from shopify all data, so that when i jmp to skipimport i have some data to paste into the daily data section
            Sheets("Shopify All Data").Range("A2:X1000000").SpecialCells(xlCellTypeVisible).Copy
            GoTo SkipImport
        Else
            Sheets("Shopify All Data").visible = False
            Sheets("Daily Data").visible = False
            Sheets("Packing Slips").Range("Q7").ClearContents
            MsgBox "Order not found"
            Exit Sub
        End If
    End If

    'Imports data from csv to initial paste page
    Call ImportCSV
    
    Dim duplicatesFound As Boolean
    
    'Call CheckForDuplicateOrder
    
    duplicatesFound = CheckForDuplicates()
    
    ' Use the boolean value as needed
    If duplicatesFound Then
        MsgBox "One or more of the orders has already been printed."
        Exit Sub
    End If
    
    'Deletes row 2 if not a valid order
    Dim ws As Worksheet
    Dim cell As Range
    Dim pattern As String
    Dim lastRowe As Long
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("Initial Paste Area")
    
    pattern = "S#####"
    
    lastRowe = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Set rng = ws.Range("A2:A" & lastRowe)
    
    For Each cell In rng
        If Not cell.value Like pattern Then
            cell.EntireRow.ClearContents
        End If
    Next cell

    Sheets("Initial Paste Area").visible = True
    Sheets("Initial Paste Area").Select
    Sheets("Initial Paste Area").AutoFilterMode = False

    'Deletes the columns that are unimportant, keeping the notes
    Range("C:H,M:O,T:T,V:AH,AR:AR,AT:BW").Select
    Selection.Delete Shift:=xlToLeft

    ' Sorting the initial paste area
    Cells.Select
    ActiveWorkbook.Worksheets("Initial Paste Area").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Initial Paste Area").Sort.SortFields.Add2 key:=Range( _
        "A:A"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Initial Paste Area").Sort.SortFields.Add2 key:=Range( _
        "L:L"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Initial Paste Area").Sort
        .SetRange Range("A:U")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A2:V10000").Select
    Selection.Copy

    Sheets("Shopify All Data").Select
    Range("A1000000").Select
    Selection.End(xlUp).Select
    activeCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
SkipImport:
        
    ' Pasting into Daily Data
    Sheets("Daily Data").Select
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    If Range("B3").value <> "" Then
        'puts in envelope group (flower, lettuce, etc.) into column A
        Range("A2").Select
        activeCell.FormulaR1C1 = _
            "=IF(RC[11]="""","""",XLOOKUP(RC[11],PLANTSKU,PLANTTYPE))"
        Range("A2").AutoFill Destination:=Range("A2:A" & Cells(Rows.Count, "B").End(xlUp).row)

        'puts in shelf location into column W
        Range("W2").Select
        activeCell.FormulaR1C1 = _
            "=IF(RC[-11]="""","""",XLOOKUP(RC[-11],PLANTSKU,RACKLOC))"
        Range("W2").AutoFill Destination:=Range("W2:W" & Cells(Rows.Count, "B").End(xlUp).row)
    Else
        Range("A2") = "None"
        Range("W2") = 1
    End If
 
    Call BulkLabelPrep

    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''
    
    Sheets("Shopify All Data").visible = False
    Sheets("Initial Paste Area").visible = False
    Sheets("Packing Slips").Select

    Dim myStart As Long
    Dim myEnd As Long
    Dim diff As Long
    'Dim lastRow As Long
    Dim i As Long
    Dim pagesToPrint As Integer
    
    myStart = CLng(Right(Sheets("Daily Data").Range("B2"), 5))
    myEnd = Range("M4").value
    
    Sheets("Packing Slips").Unprotect
    Sheets("Packing Data").visible = True
    Sheets("Intermediate").visible = True
    
    '''********* IF STATEMENT TO SEE IF THE ONEOFF ORDER NUMBER IS LESS THAN myStart.
    '  If so, it should go to the shopify all data and filter for that order (maybe call a function or sub)
        
    Sheets("Daily Data").Select

    '---------- creating a list of bulk orders ----------
    Dim bulkOrders As New Collection
    Dim orderNum As Long
    
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        orderNum = CLng(Right(Cells(i, "B"), 5))
        If InStr(1, ActiveSheet.Cells(i, "L").value, "pkt") = 0 Then
            If Not CollectionContainsValue(bulkOrders, orderNum) Then
                bulkOrders.Add orderNum
            End If
        End If
    Next i
    
    '---------- creating list of pkt orders ----------
    Dim packetOrders As New Collection
    
    For i = myStart To myEnd
        If Not CollectionContainsValue(bulkOrders, i) Then
            packetOrders.Add i
        End If
    Next i

    Call Sheet_Printer
    Sheets("Packing Slips").Select
    
    '---------- prints packet only orders ----------
    For Each packetOrder In packetOrders
        
        Sheets("Intermediate").Range("A1:W10000").ClearContents
        Sheets("Packing Data").Range("B2:W500").ClearContents
        Range("Q7").value = packetOrder
        
        'Filters for specific packet order
        Call FilterForOrder
        Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).Copy
        
        'Pastes the packet data on the Intermediate Page
        Sheets("Intermediate").Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        Call SortActivePage

        Sheets("Packing Slips").Select
        
        'Skips the packing slip if blank (for gift certificate only orders)
        If Range("C19").value <> "" Then
            pagesToPrint = Range("AREASELECT").value
            printPackingSlip (pagesToPrint)
        End If
        
        Call DeletePackingData
        
    Next packetOrder

    '---------- prints bulk orders ----------
    For Each bulkOrder In bulkOrders
        Sheets("Intermediate").Range("A1:W10000").ClearContents
        Sheets("Packing Data").Range("B2:W500").ClearContents
        Range("Q7").value = bulkOrder
        
        'Filters for bulk items
        Call FilterForOrder
        
        ActiveSheet.Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).AutoFilter Field:=12, Criteria1:="<>*pkt*" _
            , Operator:=xlAnd
        Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).Copy
    
        'Pastes the bulk items on the intermediate page
        Sheets("Intermediate").Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
                  
        Call SortActivePage

        'Returns to get the Packets
        Call FilterForOrder
        ActiveSheet.Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).AutoFilter Field:=12, Criteria1:="=*pkt*" _
            , Operator:=xlAnd
        Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).Copy
        
        'Pastes the packet data on the Intermediate Page
        Sheets("Intermediate").Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    
        Call SortActivePage
        
        'Goes to Packing Slip to print selected print areas
        Sheets("Packing Slips").Select
        
        'Skips the packing slip if blank (for gift certificate only orders)
        If Range("C19").value <> "" Then
            pagesToPrint = Range("AREASELECT").value
            printPackingSlip (pagesToPrint)
        End If
    
        Call DeletePackingData
    Next bulkOrder
    
    'records last order printed
    Range("M4").Copy
    Range("AO1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Sheets("Daily Data").AutoFilterMode = False
    
    If myStart <> myEnd Then
        Call CombineOrders
    End If
    
    Sheets("Daily Data").Range("A2:W100000").ClearContents
    Sheets("Daily Data").visible = False
    Sheets("Seed Data").visible = False
    Sheets("Packing Data").visible = False
    Sheets("Intermediate").visible = False
    Sheets("Initial Paste Area").Unprotect
    Sheets("Initial Paste Area").Range("A1:W100000").ClearContents
    Sheets("Packing Slips").Select
    Range("P5:P6").ClearContents
    Sheets("Packing Slips").Protect
    Range("L1").Select
    
    If Range("ONEOFF").value <> "" Then
        Call ViewSinglePackingSlip
    End If
    
    If Range("ONEOFF").value = "" Then
        Sheets("Packing Slips").Range("Q7").ClearContents
    End If
    
    Sheets("Packing Slips").Range("P23:P24").ClearContents
    Sheets("Packing Slips").Range("Q7").ClearContents
    Sheets("Packing Data").Range("B2:W1000").ClearContents
    
  If Sheets("Single Labels").Range("G13") <> "" Or Sheets("Single Labels").Range("R17") <> "" Then
        Sheets("Single Labels").visible = True
        Sheets("Single Labels").Select
        Range("B12").Select
        ActiveWindow.ScrollRow = 1
    End If
    
    Application.ScreenUpdating = True
    
End Sub
Sub DeletePackingData()

    Sheets("Packing Data").Range("B2:W10000").ClearContents
    
End Sub
Sub FilterForOrder()

    Sheets("Daily Data").Select
    ActiveSheet.AutoFilterMode = False
    Range("A1:W1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).AutoFilter Field:=2, Criteria1:=Range("LISTFILTER").value
    
End Sub

Function CheckForDuplicates() As Boolean
    Dim initialPasteSheet As Worksheet
    Dim shopifyDataSheet As Worksheet
    Dim initialPasteRange As Range
    Dim shopifyRange As Range
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim value As Variant
    Dim duplicateFound As Boolean
    
    ' Set reference to the worksheets
    Set initialPasteSheet = ThisWorkbook.Sheets("Initial Paste Area")
    Set shopifyDataSheet = ThisWorkbook.Sheets("Shopify All Data")
    
    ' Get the range of data from "Initial Paste Area"
    Set initialPasteRange = initialPasteSheet.Range("A2", initialPasteSheet.Range("A2").End(xlDown))
    
    ' Initialize a collection to store unique values
    Set uniqueValues = New Collection
    
    ' Loop through the range and add unique values to the collection
    For Each cell In initialPasteRange
        If cell.value <> "" Then
            On Error Resume Next ' Avoid adding duplicates to the collection
            uniqueValues.Add cell.value, CStr(cell.value)
            On Error GoTo 0
        End If
    Next cell
    
    ' Loop through the unique values and check if they exist in "Shopify All Data"
    For Each value In uniqueValues
        Set shopifyRange = shopifyDataSheet.Columns("A").Find(value, LookIn:=xlValues, LookAt:=xlWhole)
        If Not shopifyRange Is Nothing Then
            duplicateFound = True
            Exit For
        End If
    Next value
    
    ' Return the boolean value indicating whether duplicates were found
    CheckForDuplicates = duplicateFound
End Function

Function FilterForOldOrder() As Boolean
    FilterForOldOrder = False

    ThisWorkbook.Sheets("Shopify All Data").Range("A1:X100000").AutoFilter Field:=1, Criteria1:=Sheets("Packing Slips").Range("LISTFILTER").value
    
    Dim listFilter As Range
    Set listFilter = Sheets("Packing Slips").Range("LISTFILTER")
    
    Dim visible As Range
    'On Error Resume Next
    Set visible = Sheets("Shopify All Data").Range("A:A").SpecialCells(xlCellTypeVisible)

    Dim result As Range
    Set result = visible.Find(What:=listFilter.value, LookIn:=xlValues, LookAt:=xlWhole)
    
    FilterForOldOrder = Not result Is Nothing

End Function
Sub SortActivePage()
    
    'Sorts data by Qty, Shelf Location, Group and SKU (Columns I, A, & L)
    Cells.Select
    ActiveWorkbook.Worksheets("Intermediate").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Intermediate").Sort.SortFields.Add2 key:=Range( _
        "I2:I300"), SortOn:=xlSortOnValues, order:=xlDescending, DataOption:= _
        xlSortNormal
        
    ActiveWorkbook.Worksheets("Intermediate").Sort.SortFields.Add2 key:=Range( _
        "W2:W300"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal

    ActiveWorkbook.Worksheets("Intermediate").Sort.SortFields.Add2 key:=Range( _
        "L2:L300"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Intermediate").Sort
        .SetRange Range("A1:W300")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("B2:W300").Copy

    'Copies the packet information to the Packing Data Page
    Sheets("Packing Data").Select
    Range("B10000").Select
    Selection.End(xlUp).Select
    
    'This if block will put a blank line between bulk and packet items on the packing slip
    If activeCell.Address = "$B$1" Then
        Selection.Offset(1, 0).Select
    Else
        Selection.Offset(2, 0).Select
    End If
    
    ActiveSheet.Paste
    Application.CutCopyMode = False
          
    'Clears the Intermediate Page
    Sheets("Intermediate").Range("A1:W10000").ClearContents
    
End Sub
Sub PrintOldOrder()
    'this needs to pull an old order from 'shopify all data' and load it into the daily data page
    
    Call DeletePackingData
    Call FilterForOldOrder
    
    ActiveSheet.Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).AutoFilter Field:=12, Criteria1:="<>*pkt*" _
            , Operator:=xlAnd
    Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).Copy
    
    'Pastes the bulk items on the intermediate page
    Sheets("Intermediate").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Call SortActivePage
    
    Call FilterForOldOrder
    ActiveSheet.Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).AutoFilter Field:=12, Criteria1:="=*pkt*" _
            , Operator:=xlAnd
    Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).Copy
    
    'Pastes the packet data on the Intermediate Page
    Sheets("Intermediate").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        
    Call SortActivePage
    
    'Goes to Packing Slip to print selected print areas
    Sheets("Packing Slips").Select
    'Skips the packing slip if blank (for gift certificate only orders)
    If Range("C19").value <> "" Then
            pagesToPrint = Range("AREASELECT").value
            printPackingSlip (pagesToPrint)
    End If

    Range("P6:P7").ClearContents
        
    Sheets("Shopify All Data").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    Range("A2").Select
    
End Sub
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

'Function that checks to see if a value is already in the list
Function CollectionContainsValue(col As Collection, val As Variant) As Boolean
    Dim n As Long
    CollectionContainsValue = False
    For n = 1 To col.Count
        If col.Item(n) = val Then
            CollectionContainsValue = True
            Exit Function
        End If
    Next n
End Function

'Print function
Function printPackingSlip(pagesToPrint As Integer)
    
    If pagesToPrint = 1 Then GoTo AreaOne
    If pagesToPrint = 2 Then GoTo AreaTwo
    If pagesToPrint = 3 Then GoTo AreaThree
    If pagesToPrint = 4 Then GoTo AreaFour
    If pagesToPrint = 5 Then GoTo AreaFive
    If pagesToPrint = 6 Then GoTo AreaSix
    If pagesToPrint = 7 Then GoTo AreaSeven
    If pagesToPrint = 8 Then GoTo AreaEight

AreaOne:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$41"
    GoTo PrintPage
        
AreaTwo:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$78"
    GoTo PrintPage
            
AreaThree:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$115"
    GoTo PrintPage
            
AreaFour:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$152"
    GoTo PrintPage
            
AreaFive:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$189"
    GoTo PrintPage
                
AreaSix:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$226"
    GoTo PrintPage
    
AreaSeven:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$263"
    GoTo PrintPage
    
AreaEight:
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$300"
    
PrintPage:

    ActiveWindow.SelectedSheets.PrintOut copies:=Range("PACKSLIPQTY").value, Collate:=True, _
        IgnorePrintAreas:=False
    
End Function

Sub BulkLabelPrep()
    Dim seedData As Worksheet
    Dim sourceRng As Range
    Dim rngVisible As Range
    Dim lastRow As Long
    Dim cell As Range
    Dim crit As Variant
    
    Sheets("Single Labels").Unprotect
    Sheets("Single Labels").Range("G13:H600").ClearContents
    Sheets("Single Labels").Range("R17:U106").ClearContents
    Sheets("Single Labels").Range("B12:B13").ClearContents
    Sheets("Seed Data").Unprotect
    Sheets("Seed Data").visible = True
    Sheets("Seed Data").Select
    
    Rows("1:1").Select
    Selection.AutoFilter
    
    Set seedData = ThisWorkbook.Sheets("Seed Data")
    
    seedData.Range("$A$1:$BJ$1501").AutoFilter Field:=1, Criteria1:="<>*Pkt*" _
        , Operator:=xlAnd

    seedData.Range("$A$1:$BJ$1501").AutoFilter Field:=61, Criteria1:=">0", _
        Operator:=xlAnd
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim toPrintDict As Object
    Dim toPackDict As Object
    Dim sku As String
    Dim qtyOrdered As Double
    Dim qtyPrePack As Double
    Dim selectedCell As Range
    
    Set toPrintDict = CreateObject("Scripting.Dictionary")
    Set toPullDict = CreateObject("Scripting.Dictionary")
    
    
    Set sourceRng = seedData.Range("A2:BI1501")
    
    
    On Error Resume Next
    If Not sourceRng.SpecialCells(xlCellTypeVisible) Is Nothing Then
       For Each cell In sourceRng.Columns(1).SpecialCells(xlCellTypeVisible)

        sku = cell.value
        qtyOrdered = cell.Offset(0, 60).value
        qtyPrePack = cell.Offset(0, 61).value
        
        If qtyPrePack > 0 Then
            If qtyPrePack >= qtyOrdered Then
                Dim cellVal As Integer
                cellVal = cell.Offset(0, 61).value
                cell.Offset(0, 61).value = qtyPrePack - qtyOrdered
                toPullDict(sku) = qtyOrdered
                cellVal = cell.Offset(0, 61).value
                GoTo ContinueLoop
            Else
                toPullDict(sku) = qtyPrePack
                toPrintDict(sku) = qtyOrdered - qtyPrePack
                cell.Offset(0, 61).value = 0
                GoTo ContinueLoop
            End If
                
        Else
            toPrintDict(sku) = qtyOrdered
        End If
            
            
ContinueLoop:
          
        Next cell
        
        Dim singleLabels As Worksheet
        Set singleLabels = ThisWorkbook.Sheets("Single Labels")
        
        ' Inserting data from dictionaries onto Single Labels page
        Dim keyTwo As Variant
        Dim row As Long
        row = 13
        
        For Each keyTwo In toPrintDict.Keys
            If Not (InStr(1, keyTwo, "MER", vbTextCompare) > 0 And _
                    InStr(1, keyTwo, "BOOK", vbTextCompare) > 0 And _
                    InStr(1, keyTwo, "TOO", vbTextCompare) > 0 And _
                    InStr(1, keyTwo, "SKU", vbTextCompare) > 0 And _
                    InStr(1, keyTwo, "gift", vbTextCompare) > 0) Then
                    
                singleLabels.Cells(row, "G").value = keyTwo
                singleLabels.Cells(row, "H").value = toPrintDict(keyTwo)
                row = row + 1
            End If
        Next keyTwo
        
        row = 17
        
        For Each keyTwo In toPullDict.Keys
            singleLabels.Cells(row, "R").value = keyTwo
            singleLabels.Cells(row, "T").value = toPullDict(keyTwo)
            row = row + 1
        Next keyTwo
        
        ' Checks to see if there are more than one bulk item. If not, skips the sorting process
        If Not IsEmpty(singleLabels.Range("G14").value) Then

            Dim lastRowTwo As Long
            Dim sortRange As Range
        
            lastRowTwo = singleLabels.Cells(singleLabels.Rows.Count, "G").End(xlUp).row
            Set sortRange = singleLabels.Range("G13:I" & lastRowTwo)
            sortRange.Sort Key1:=sortRange.Cells(1, 3), Order1:=xlAscending, Header:=xlNo, Key2:=sortRange.Cells(1, 1), Order2:=xlAscending
              
        End If
        
        If singleLabels.Range("H13") = 0 Then
            singleLabels.Range("H13").ClearContents
        End If
        
    End If
    
    seedData.AutoFilterMode = False
  
End Sub
Function HasPartialMatch(value As String, criteria As Variant) As Boolean
    Dim crit As Variant
    For Each crit In criteria
        If InStr(1, value, crit, vbTextCompare) > 0 Then
            HasPartialMatch = True
            Exit Function
        End If
    Next crit
    
    HasPartialMatch = False

End Function
Sub PrintSinglePackingSlip()

    Call PrintPackingSlips
    
End Sub
Sub ViewSinglePackingSlip()
     Application.ScreenUpdating = False
    
    If Range("ONEOFF").value = "" Then
        MsgBox "Please Enter an order number"
        Exit Sub
    End If

    Dim packingSlip As Worksheet
    Dim shopifyAll As Worksheet
    Dim intermediate As Worksheet
    Dim dailyData As Worksheet
    Dim packingData As Worksheet
    Dim orderNum As String
    Dim lastRow As Long
    Dim order As Range
    Dim formulaRange As Range
    Dim bulk As Range
    Dim packet As Range
    
    Set packingSlip = ThisWorkbook.Sheets("Packing Slips")
    Set shopifyAll = ThisWorkbook.Sheets("Shopify All Data")
    Set intermediate = ThisWorkbook.Sheets("Intermediate")
    Set dailyData = ThisWorkbook.Sheets("Daily Data")
    Set packingData = ThisWorkbook.Sheets("Packing Data")
    
    packingData.Unprotect
    packingData.Range("B2:W1000").ClearContents
    dailyData.Unprotect
    dailyData.Range("A2:W1000").ClearContents
    intermediate.Unprotect
    intermediate.AutoFilterMode = False
    intermediate.Range("A2:W1000").ClearContents

    orderNum = "S" & Right(packingSlip.Range("ONEOFF").value, 5)
    
    shopifyAll.Unprotect
    shopifyAll.AutoFilterMode = False
    lastRow = shopifyAll.Cells(shopifyAll.Rows.Count, "A").End(xlUp).row
    
    shopifyAll.Range("A1:A" & lastRow).AutoFilter Field:=1, Criteria1:=orderNum
    Set order = shopifyAll.UsedRange.SpecialCells(xlCellTypeVisible)
    
   ' Exit Sub
    
    If Not order Is Nothing Then
        order.Copy dailyData.Range("B1")
        order.Copy intermediate.Range("A1")
        lastRow = intermediate.Cells(intermediate.Rows.Count, "A").End(xlUp).row
                
        intermediate.Range("V1").Formula = "=XLOOKUP(K1,PLANTSKU,RACKLOC)"
        
        Sheets("Packing Slips").Range("Q7").value = orderNum
        
        Set formulaRange = intermediate.Range("V1:V" & lastRow)
        
        On Error GoTo ErrorHandler:
        intermediate.Range("V1").AutoFill Destination:=formulaRange
        On Error GoTo 0
        
        With intermediate.Sort
            .SortFields.Clear
            .SortFields.Add key:=intermediate.Range("H1:H" & lastRow), SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortNormal
            .SortFields.Add key:=intermediate.Range("V1:V" & lastRow), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .SetRange intermediate.Range("A1:V" & lastRow)
            .Header = xlNo
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        intermediate.AutoFilterMode = False
        
        intermediate.Range("K1:K" & lastRow).AutoFilter Field:=1, Criteria1:="<>*pkt*", Operator:=xlAnd
        
        Set bulk = intermediate.UsedRange.Offset(1).SpecialCells(xlCellTypeVisible)
        
        If Not bulk Is Nothing Then
            'MsgBox "inside If Not bulk Is Nothing Then"
            bulk.Copy packingData.Range("B2")
        End If
        
        intermediate.AutoFilterMode = False
        
        intermediate.Range("K1:K" & lastRow).AutoFilter Field:=1, Criteria1:="=*pkt*", Operator:=xlAnd
        
        Set packets = intermediate.UsedRange.Offset(1).SpecialCells(xlCellTypeVisible)
        
        If Not packets Is Nothing Then
            If packingData.Range("B2").value = "" Then
                packets.Copy packingData.Range("B2")
            Else
                lastRow = packingData.Cells(packingData.Rows.Count, "B").End(xlUp).row
                packets.Copy packingData.Range("B" & lastRow + 2)
            End If
            
        End If
    Else
        MsgBox "Not valid"
    End If
    
    shopifyAll.AutoFilterMode = False
    
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    MsgBox "Order number not found"
    shopifyAll.AutoFilterMode = False
    intermediate.AutoFilterMode = False

End Sub
