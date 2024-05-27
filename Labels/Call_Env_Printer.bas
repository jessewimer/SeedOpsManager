Sub Env_Printer()

  For i = 0 To 12
  
  curNePrint = Format(i, "00")
  On Error Resume Next
  
  Application.ActivePrinter = Range("ENVPRINTER").value & " on Ne" & curNePrint & ":"
  Next i

End Sub
