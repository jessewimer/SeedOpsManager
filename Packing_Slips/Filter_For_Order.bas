Sub FilterForOrder()

    Sheets("Daily Data").Select
    ActiveSheet.AutoFilterMode = False
    Range("A1:W1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:" & "W" & Range("DAILYCOUNT").value + 3).AutoFilter Field:=2, Criteria1:=Range("LISTFILTER").value
    
End Sub
