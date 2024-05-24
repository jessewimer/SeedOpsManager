Sub QLPrintBackSingle()

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

    Application.ScreenUpdating = False
    
    Call Roll_Printer
    
    If Range("QLSKIPBACK").value = 1 Then GoTo ContinueMacro
    
    Output = MsgBox("There are no back labels to print or that option is set to NO in the SEED DATA Page", vbExclamation, "Label Data Unavailable")
    GoTo Bottom
    
ContinueMacro:
        
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
    
    Sheets("Home").Select

    Call HideAllLabels
    
Bottom:
    
    Application.ScreenUpdating = True

End Sub
