Sub PrintEnvelopeBack()

    'checks to see if back labels are set to print OR if force back printing is activated
    If Range("QLSKIPBACK").value <> 1 Then
        Output = MsgBox("There are no back labels to print or that option is set to NO in the SEED DATA Page", vbExclamation, "Label Data Unavailable")
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Call Env_Printer
    
    Sheets("Envelope Back 1").visible = True

    Sheets("Envelope Back 1").Select

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=Range("ENVPRQTY").value, Collate:=True, _
        IgnorePrintAreas:=False
        
    Sheets("Home").Select
    
    Range("B4").Select
    
    Sheets("Envelope Back 1").visible = False
    'Sheets("Envelope Back 3").Visible = False

    Application.ScreenUpdating = True

End Sub
