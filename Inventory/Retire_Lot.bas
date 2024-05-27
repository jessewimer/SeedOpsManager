Sub RecordLotChange()

    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
        
    Dim lotRange As Range
    Dim activeRange As Range
    Dim targetCell As Range
    Dim packStatus As VbMsgBoxResult
    Dim weightInput As Variant
    Dim noteInput As Variant
    
    'Set the range
    Set lotRange = Range("L21:L23")
 
    'Get the cell or range that the user selected
    Set activeRange = activeCell
    
    'Check if the selection is inside the range
    If Intersect(lotRange, activeRange) Is Nothing Then
        'Selection is NOT inside the range
        MsgBox "Invalid selection"
        
    ' Else selection is valid
    Else
        
        Sheets("Retired Lots").Unprotect
        Dim message As String
        Dim skuPrefix As Range
        Set skuPrefix = Range("S22")
        
        'MsgBox activeRange.Address
        If activeRange.MergeArea.Address = Range("L21").MergeArea.Address Then
            message = "Is this the correct lot to be retired?" & vbNewLine & vbNewLine & Range("S17").value & vbNewLine & Range("S26").value
            Set lotRange = Range("S26")
        ElseIf activeRange.MergeArea.Address = Range("L22").MergeArea.Address Then
            message = "Is this the correct lot to be retired?" & vbNewLine & vbNewLine & Range("S17").value & vbNewLine & Range("S30").value
            Set lotRange = Range("S30")
        Else
            message = "Is this the correct lot to be retired?" & vbNewLine & vbNewLine & Range("S17").value & vbNewLine & Range("S30").value
            Set lotRange = Range("S34")
        End If
        
        Dim result As VbMsgBoxResult
        result = MsgBox(message, vbOKCancel + vbInformation, "Pop-Up Message")
    
        If result = vbCancel Then
            'MsgBox "exiting sub"
            Exit Sub
        End If
        
        ' Assemble the final sku/lot to be recorded
        Dim finalSKU As String
        finalSKU = skuPrefix.value & "-" & lotRange.value
        
        ' Find 1st blank cell in Column B in "Retired Lots" sheet
        On Error Resume Next
        Set targetCell = Sheets("Retired Lots").Columns("B").Find("*", , xlValues, xlWhole, xlByColumns, xlPrevious)
        On Error GoTo 0
        
        ' Recording the final sku/lot
        Set targetCell = targetCell.Offset(1, 0)
        targetCell.value = finalSKU
        
        ' Recording the date
        Set targetCell = targetCell.Offset(0, 1)
        targetCell.value = Date
        
        ' Offset to record packing status (whether or not the lot was packed out fully)
        Set targetCell = targetCell.Offset(0, 1)
        
        ' Determing if the lot was packed all the way out
        packStatus = MsgBox("Was the lot packed all the way out?", vbYesNo + vbQuestion, "Packing Status")
        
        If packStatus = vbYes Then
            targetCell.value = 1
            Set targetCell = targetCell.Offset(0, 1)
            targetCell.value = 0
        Else
            targetCell.value = 2
            ' Offsetting to LBS column
            Set targetCell = targetCell.Offset(0, 1)
            ' Prompting the user for the LBS
            weightInput = InputBox("Record remaining weight (lbs):", "Weight Input")
            
            If IsNumeric(weightInput) Then
                targetCell.value = weightInput
            Else
                MsgBox "Not a valid weight, please try again"
                Set targetCell = targetCell.Offset(0, -1)
                targetCell.value = ""
                Set targetCell = targetCell.Offset(0, -1)
                targetCell.value = ""
                Set targetCell = targetCell.Offset(0, -1)
                targetCell.value = ""
                Exit Sub
            End If
         
        End If
        
        ' Offsetting to Notes column
        Set targetCell = targetCell.Offset(0, 1)
        
        ' Prompting user for the notes
        noteInput = InputBox("Record any applicable notes:", "Note Input")
        targetCell.value = noteInput
        
    End If
    
    Sheets("Retired Lots").Protect
    
    Application.ScreenUpdating = True

End Sub
