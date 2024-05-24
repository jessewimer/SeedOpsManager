Sub QLPrintFrontSheet()

    Dim wrkshit As Worksheet
    Dim seerchVal As Variant
    Dim sell As Range
    Dim found As Boolean
    
    ' Set the worksheet
    Set wrkshit = ThisWorkbook.Sheets("Home") ' Replace "Sheet1" with your sheet's name
    
    ' Set the value to search for
    seerchVal = wrkshit.Range("B1").value
    
    ' Initialize found flag
    found = False
    
    ' Loop through each sell in the specified range
    For Each sell In wrkshit.Range("A9:A27")
        If sell.value = seerchVal Then
            found = True
            Exit For
        End If
    Next sell
    
    ' Check if the value was found
    If Not found Then
        MsgBox "populate item plz.", vbExclamation
        Exit Sub
    End If

    If Range("K27").value = 1 Then
        Output = MsgBox("Lot or Germ not detected", vbExclamation, "Error")
        Exit Sub
    End If

    'Check to see if the lot has been retired
    
    If ActiveSheet.Range("S28").value <> "" Then
        If ActiveSheet.Range("U25").value = True Then
            MsgBox "This lot is retired."
            Exit Sub
        End If
    ElseIf ActiveSheet.Range("S32").value <> "" Then
        If ActiveSheet.Range("U29").value = True Then
            MsgBox "This lot is retired."
            Exit Sub
        End If
    ElseIf ActiveSheet.Range("S36").value <> "" Then
        If ActiveSheet.Range("U33").value = True Then
            MsgBox "This lot is retired."
            Exit Sub
        End If
    End If

    If Range("S63") <> 1 Then
        CarryOn3 = MsgBox("You are printing a full sheet of a bulk item. Do you want to continue?", vbYesNo, "Continue")
        If CarryOn3 = vbNo Then Exit Sub
        Cancel = True
    End If
    
    If Range("S61") = 1 Then
        If Range("W13") = 1 Then
            CarryOn3 = MsgBox("Low inventory. Do you want to print anyway?", vbYesNo, "Continue")
            If CarryOn3 = vbNo Then Exit Sub
        Cancel = True
        End If
    ElseIf Range("S61") = 2 Then
        If Range("W14") = 1 Then
            CarryOn3 = MsgBox("Low inventory. Do you want to print anyway?", vbYesNo, "Continue")
            If CarryOn3 = vbNo Then Exit Sub
        Cancel = True
        End If
    ElseIf Range("S61") = 3 Then
        If Range("W15") = 1 Then
            CarryOn3 = MsgBox("Low inventory. Do you want to print anyway?", vbYesNo, "Continue")
            If CarryOn3 = vbNo Then Exit Sub
        Cancel = True
        End If
    End If

    If Range("K19") = 0 Then
        myMsgBox = MsgBox("This was already printed today. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
    End If
    
    If Range("K19") = 1 Then
        myMsgBox = MsgBox("This was printed yesterday. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
    End If
    
    If Range("K19") = 2 Then
        myMsgBox = MsgBox("This was printed two days ago. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
    End If

    If Range("K19") = 3 Then
        myMsgBox = MsgBox("This was printed three days ago. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
    End If
    
    If Range("K19") > 3 And Range("K19") < 8 Then
        myMsgBox = MsgBox("This was printed within the last week. Do you wish to contine?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
    End If
    
    Application.ScreenUpdating = False
    
    'if 'pkt' in sku, records last print date/qty
    If Range("S63") > 0 Then
        Sheets("Germination Data").Unprotect
        Sheets("Germination Data").visible = True
        Sheets("Germination Data").Select
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    
        Dim cell As Range
        Set cell = Range("A:A").Find(What:=Range("CE1").value, LookIn:=xlValues, LookAt:=xlWhole)
    
        If cell Is Nothing Then
            Output = MsgBox("Please enter SKU into cell B1 on the Home page", vbExclamation, "Error")
        
        Else
            cell.Select
            activeCell.Offset(, 72).Select
            'MsgBox activeCell.Value
            'Exit Sub
            
            'increments total number of pkts printed
            activeCell.value = activeCell.value + (Range("BX1").value * 30)
        
            'offsetting one column to the right to record how many were printed on the last print date
            activeCell.Offset(, 2).Select
            
            Dim adjacentCell As Range
            Set adjacentCell = activeCell.Offset(0, -1)
            
            If adjacentCell.value = Date Then
                activeCell.value = activeCell.value + (Range("BX1").value * 30)
            Else
                activeCell.value = (Range("BX1").value * 30)
                
            End If
        
            'offsetting one column to the right to record last print date
            activeCell.Offset(, -1).Select
            activeCell.value = Date
            activeCell.Copy
            activeCell.PasteSpecial Paste:=xlPasteValues
            
        End If
    End If

    Call Sheet_Printer
        
    Call UnhideAllLabels
    
    Call QLCopyB1
    
    If Range("QLFRONTLABNUM").value = 0 Then GoTo Bottom
    
    If Range("S63") = 0 Then GoTo PrintBulk
    
    If Range("QLFRONTLABNUM").value = 1 Then GoTo PrintSet1
     
        Sheets("Label 2").Select
        Range("O8:P9").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        GoTo PrintLabels
        
PrintSet1:
        Sheets("Label 1").Select
        Range("O8:P9").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        GoTo PrintLabels
        
PrintBulk:
        Sheets("Home").Select
        If Range("Q1").value = 1 Then
            Sheets("Bulk Sheet (3)").Select
            Sheets("Bulk Sheet (3)").Unprotect
        Else
            Sheets("Bulk Sheet").Select
            Sheets("Bulk Sheet").Unprotect
            
        End If
    
PrintLabels:
        ActiveWindow.SelectedSheets.PrintOut copies:=Range("QLPRTCP").value, Collate:=True, _
            IgnorePrintAreas:=False
            
        Range("O8:P9").Select
        Selection.ClearContents
        
        Sheets("Home").Select
        
Bottom:
    
    Call HideAllLabels
    
    Application.ScreenUpdating = True

End Sub
