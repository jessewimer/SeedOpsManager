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
