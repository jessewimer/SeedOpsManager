Sub QLPrintFrontBackSingle()

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
    
    Application.ScreenUpdating = False

    'if 'pkt' in sku, records last print date/qty
    If Range("S63").value > 0 Then
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
            
            'increments total number of pkts printed
            activeCell.value = activeCell.value + (Range("BX1").value)
        
            'offsetting one column to the right to record how many were printed on the last print date
            activeCell.Offset(, 2).Select
            
            Dim adjacentCell As Range
            Set adjacentCell = activeCell.Offset(0, -1)
            
            If adjacentCell.value = Date Then
                activeCell.value = activeCell.value + (Range("BX1").value)
            Else
                activeCell.value = (Range("BX1").value)
                
            End If
        
            'offsetting one column to the right to record last print date
            activeCell.Offset(, -1).Select
            activeCell.value = Date
            activeCell.Copy
            activeCell.PasteSpecial Paste:=xlPasteValues
            
        End If
    End If


    Call Roll_Printer
    
    
    If Range("QLSKIPBACK").value = 2 Then GoTo PrintFrontOnly
    
    Call UnhideAllLabels
     
    If Range("QLBACKNUM").value = 7 Then GoTo PrintSingle1
 
    Sheets("Back Label 3").Select
    GoTo PrintLabels
 
PrintSingle1:
    Sheets("Back Label 1").Select
    GoTo PrintLabels

PrintLabels:
    
    ActiveWindow.SelectedSheets.PrintOut copies:=Range("QLPRTCP").value, Collate:=True, _
        IgnorePrintAreas:=False
    
    'Sheets("Home").Select
    
PrintFrontOnly:
    'MsgBox "inside printfrontonly"
    Call UnhideAllLabels
    
    'MsgBox QLFRONTLABNUM
    
    Dim cellVal As Variant
    cellVal = Range("QLFRONTLABNUM").value
    
    'If Range("QLFRONTLABNUM").value = 0 Then GoTo Bottom
    If Range("QLFRONTLABNUM").value = 1 Then GoTo PrintSet1
    
    If Sheets("Home").Range("S63").value > 0 Then
        Sheets("Single Label 2").Select
    Else
        If Sheets("Home").Range("S65").value = "" Then
        
            Sheets("Bulk Label Template").visible = True
            Sheets("Bulk Label Template").Select
            
        Else
            Sheets("Bulk Label Template Radicchio").visible = True
            Sheets("Bulk Label Template Radicchio").Select
        End If
    End If
    GoTo PrintSingles
    
PrintSet1:
    'if pkt in sku
    If Sheets("Home").Range("S63").value > 0 Then
        Sheets("Single Label 1").Select
    Else
        Sheets("Bulk Label Template 2").Select
    End If

PrintSingles:
    'Sheets("Single Label 1").Select
    ActiveWindow.SelectedSheets.PrintOut copies:=Range("QLPRTCP").value, Collate:=True, _
        IgnorePrintAreas:=False

    Sheets("Home").Select

    If Range("QLSKIPBACK").value = 1 Then GoTo Bottom
    
    Output = MsgBox("There were no back labels to print or that option is set to NO in the SEED DATA Page", vbExclamation, "Label Data Unavailable")
    
Bottom:
        
    Call HideAllLabels
    Application.ScreenUpdating = True

End Sub
