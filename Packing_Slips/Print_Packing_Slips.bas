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
