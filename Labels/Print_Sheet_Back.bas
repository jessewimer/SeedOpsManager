Sub QLPrintBackSheet()

    Dim wrkshit As Worksheet
    Dim seerchVal As Variant
    Dim sell As Range
    Dim found As Boolean
    
    ' Set the worksheet
    Set wrkshit = ThisWorkbook.Sheets("Home") ' Replace "Sheet1" with your sheet's name
    
    ' Set the value to search for
    seerchVal = wrkshit.Range("S23").value
    
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

    Application.ScreenUpdating = False
    
    Call Sheet_Printer
    
    If Range("QLSKIPBACK").value = 1 Then GoTo ContinueMacro
    
    Output = MsgBox("There are no back labels to print or that option is set to NO in the SEED DATA Page", vbExclamation, "Label Data Unavailable")
    
    GoTo Bottom
    
ContinueMacro:
        
    Call UnhideAllLabels
 
    Call QLCopyB1
 
    If Range("QLBACKNUM").value = 7 Then GoTo PrintSheet1
 
    Sheets("Back Label Sheet 3").Select
    Range("Q19:R20").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    GoTo PrintLabels
 
PrintSheet1:
    Sheets("Back Label Sheet 1").Select
    Range("Q19:R20").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    GoTo PrintLabels
    
PrintLabels:
    
    ActiveWindow.SelectedSheets.PrintOut copies:=Range("QLPRTCP").value, Collate:=True, _
        IgnorePrintAreas:=False
        
    Range("Q19:R20").Select
    Selection.ClearContents
    
    Sheets("Home").Select

    Call HideAllLabels
    
Bottom:
    
    Application.ScreenUpdating = True

End Sub
