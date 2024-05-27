Sub AddNewLot()
    
    Application.ScreenUpdating = False
    
    Sheets("Germination Data").Unprotect

    Dim lotRange As Range
    Set lotRange = Range("L21:L23")
    
    If Not Intersect(activeCell, lotRange) Is Nothing Then
        If activeCell = Range("L21") Then
            Range("'Germination Data'!CD1") = 1
            If Range("L21").value <> "" Then
                    
                Dim response1 As VbMsgBoxResult
                Dim newLotNumber1 As String
            
                ' Prompt the user with a message box
                response1 = MsgBox("Do you want to change the lot number?", vbYesNo + vbQuestion, "Lot Number Change")
            
                ' Check the user's response
                If response1 = vbYes Then
                    ' If the user clicks "Yes", prompt for the new lot number
                    newLotNumber1 = InputBox("Enter the new lot number:", "New Lot Number")

                Else
                    GoTo Bottom
                End If
            Else
                newLotNumber1 = InputBox("Enter the new lot number:", "New Lot Number")
            End If
                
        ElseIf activeCell = Range("L22") Then
            Range("'Germination Data'!CD1") = 2
            If Range("L22").value <> "" Then
                    
                Dim response2 As VbMsgBoxResult
                Dim newLotNumber2 As String
            
                ' Prompt the user with a message box
                response2 = MsgBox("Do you want to change the lot number?", vbYesNo + vbQuestion, "Lot Number Change")
            
                ' Check the user's response
                If response2 = vbYes Then
                    ' If the user clicks "Yes", prompt for the new lot number
                    newLotNumber2 = InputBox("Enter the new lot number:", "New Lot Number")
            
                Else
                    GoTo Bottom
                End If
            Else
                newLotNumber2 = InputBox("Enter the new lot number:", "New Lot Number")
            End If
            
        ElseIf activeCell = Range("L23") Then
            Range("'Germination Data'!CD1") = 3
            If Range("L23").value <> "" Then
                    
                Dim response3 As VbMsgBoxResult
                Dim newLotNumber3 As String
            
                ' Prompt the user with a message box
                response3 = MsgBox("Do you want to change the lot number?", vbYesNo + vbQuestion, "Lot Number Change")
            
                ' Check the user's response
                If response3 = vbYes Then
                    ' If the user clicks "Yes", prompt for the new lot number
                    newLotNumber3 = InputBox("Enter the new lot number:", "New Lot Number")
            
                Else
                    GoTo Bottom
                End If
            Else
                newLotNumber3 = InputBox("Enter the new lot number:", "New Lot Number")
            End If
        End If
    Else
        MsgBox "Please select a lot number"
        Exit Sub
    End If

    Sheets("Germination Data").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    
    Dim cell As Range
    Set cell = Range("A:A").Find(What:=Range("CE1").value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        Output = MsgBox("Please enter SKU into cell B1 on the Home page", vbExclamation, "Error")
    Else
        cell.Select
        If Range("CD1") = 1 Then
            activeCell.Offset(, 4).value = newLotNumber1
        ElseIf Range("CD1") = 2 Then
            activeCell.Offset(, 10).value = newLotNumber2
        Else
            activeCell.Offset(, 16).value = newLotNumber3
        End If
    End If
    
Bottom:

    Sheets("Germination Data").Protect AllowFiltering:=True
    Sheets("Home").Select
    
    Application.ScreenUpdating = True
    
End Sub
