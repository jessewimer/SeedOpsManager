Sub QLChangeActiveLot()

    Application.ScreenUpdating = False

    Sheets("Seed Data").Unprotect
        
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
        
         ' Prompt the user if they want to change the active lot for all sizes
        Dim response As VbMsgBoxResult
        response = MsgBox("Do you want to change the active lot for all sizes?", vbYesNo + vbQuestion, "Change Active Lot for All Sizes")
        
        If response = vbYes Then
            ' User wants to change active lot for all sizes
            ' Capture this information in a variable
            Dim changeForAllSizesFlag As Boolean
            changeForAllSizesFlag = True
        End If
         
        Dim cellL21 As Range
        Dim cellL22 As Range
        Dim cellL23 As Range
    
        ' Set references to the cells
        Set cellL21 = Range("L21")
        Set cellL22 = Range("L22")
        Set cellL23 = Range("L23")
        
        If activeRange = cellL21 Then
            Range("'Seed Data'!CA1") = 1
        ElseIf activeRange = cellL22 Then
            Range("'Seed Data'!CA1") = 2
        Else
            Range("'Seed Data'!CA1") = 3
        End If
        
        
        Sheets("Seed Data").visible = True
        Sheets("Seed Data").Select
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    
        Dim cell As Range
        Set cell = Range("A:A").Find(What:=Range("CB1").value, LookIn:=xlValues, LookAt:=xlWhole)
    
        If cell Is Nothing Then
            Output = MsgBox("Please enter SKU into cell B1 on the Home page", vbExclamation, "Error")
        Else
        
            If changeForAllSizesFlag = True Then
                Dim cb1Value As String
                cb1Value = Left(Range("CB1").value, 6)
                
                For Each cell In Range("A2:A1500")
                    If Left(cell.value, 6) = cb1Value Then
                        cell.Select
                        activeCell.Offset(, 17).Select
                        activeCell = Range("CA1")
                    End If
                Next cell
            
            Else
            
                cell.Select
                activeCell.Offset(, 17).Select
                activeCell = Range("CA1")
                
            End If
        End If
    End If

    Sheets("Seed Data").Protect
    Sheets("Home").Select
    Application.ScreenUpdating = True

End Sub
