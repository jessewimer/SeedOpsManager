Sub PrintEnvelopeFront()
  
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
    
    If Range("K19") = 0 Then
        myMsgBox = MsgBox("This was already printed today. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
        'End If
    End If

    If Range("K19") = 1 Then
        myMsgBox = MsgBox("This was printed yesterday. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
        'End If
    End If
    
    If Range("K19") = 2 Then
        myMsgBox = MsgBox("This was printed two days ago. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
        'End If
    End If
    
    If Range("K19") = 3 Then
        myMsgBox = MsgBox("This was printed three days ago. Do you wish to continue?", vbYesNo, "Continue")
        If myMsgBox = vbNo Then Exit Sub
        Cancel = True
        'End If
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
            activeCell.value = activeCell.value + (Range("ENVPRQTY").value)
        
            'offsetting one column to the right to record how many were printed on the last print date
            activeCell.Offset(, 2).Select
            
            Dim adjacentCell As Range
            Set adjacentCell = activeCell.Offset(0, -1)
            
            If adjacentCell.value = Date Then
                activeCell.value = activeCell.value + (Range("ENVPRQTY").value)
            Else
                activeCell.value = (Range("ENVPRQTY").value)
                
            End If
        
            'offsetting one column to the right to record last print date
            activeCell.Offset(, -1).Select
            activeCell.value = Date
            activeCell.Copy
            activeCell.PasteSpecial Paste:=xlPasteValues
            
        End If
    End If
  
    Call Env_Printer
    
    If Range("QLFRONTLABNUM").value = 0 Then GoTo Bottom
     
    If Range("QLFRONTLABNUM").value = 1 Then GoTo PrintSet1
     
    Sheets("Envelope Front 2").visible = True
    Sheets("Envelope Front 2").Select
        
    GoTo PrintLabels
        
PrintSet1:
        Sheets("Envelope Front 1").visible = True
        Sheets("Envelope Front 1").Select
        
PrintLabels:
    
        ActiveWindow.SelectedSheets.PrintOut From:=1, To:=Range("ENVPRQTY").value, Collate:=True, _
            IgnorePrintAreas:=False
        Sheets("Home").Select
        Range("B5").Select
        
Bottom:
    
    Sheets("Envelope Front 1").visible = False
    Sheets("Envelope Front 2").visible = False
    
    Application.ScreenUpdating = True

End Sub
