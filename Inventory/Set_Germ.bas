Sub SetGerm()
    
    Application.ScreenUpdating = False
    
    Sheets("Germination Data").Unprotect

    Dim lotRange As Range
    Set lotRange = Range("L21:L23")
    
    If Not Intersect(activeCell, lotRange) Is Nothing Then
        If activeCell = Range("L21") Then
            Range("'Germination Data'!CD1") = 1
        ElseIf activeCell = Range("L22") Then
            Range("'Germination Data'!CD1") = 2
        ElseIf activeCell = Range("L23") Then
            Range("'Germination Data'!CD1") = 3
        End If
    Else
        MsgBox "Please select a lot number"
        Exit Sub
    End If
    
    ' get user input
    Dim lbsAmt As Variant
    ' Prompt user for input
    lbsAmt = Application.InputBox("Enter the germination rate:", Type:=1)
    
    Sheets("Germination Data").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    
    Dim cell As Range
    Set cell = Range("A:A").Find(What:=Range("CE1").value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        Output = MsgBox("Please enter SKU into cell B1 on the Home page", vbExclamation, "Error")
    Else
        cell.Select
        If Range("CD1") = 1 Then
            activeCell.Offset(, 7).value = lbsAmt
        ElseIf Range("CD1") = 2 Then
            activeCell.Offset(, 13).value = lbsAmt
        Else
            activeCell.Offset(, 19).value = lbsAmt
        End If
    End If
    
    Sheets("Germination Data").Protect AllowFiltering:=True
    Sheets("Home").Select
    
    Application.ScreenUpdating = True
    
End Sub
