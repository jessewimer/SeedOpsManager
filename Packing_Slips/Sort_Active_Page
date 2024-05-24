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
