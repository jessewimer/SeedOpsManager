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
