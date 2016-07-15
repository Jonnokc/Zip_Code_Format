Sub ZipCodeCleanUp()

Dim LR As Long
Dim Copy As String

    LR = Range("C" & Rows.Count).End(xlUp).Row
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Distance"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=R[1]C[1]-RC[1]"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & LR)
    
    LR = Range("C" & Rows.Count).End(xlUp).Select
    ActiveCell.offset(-1, -1).Select
    Copy = ActiveCell.Value
    ActiveCell.offset(1, 0).Select
    ActiveCell.Value = Copy
    
    Range("B2").Select
    Range("B1").AutoFilter Field:=1, Criteria1:=">20", _
    Operator:=xlOr, Criteria2:="<0"
        Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Bad"
    Cells.AutoFilter
    Selection.AutoFilter
    Range("B1").AutoFilter Field:=1, Operator:= _
        xlFilterNoFill
        Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Good"
    Cells.AutoFilter
    Columns("C").Select
    Selection.NumberFormat = "00000"
    Columns("B:B").Select
    Selection.NumberFormat = "General"
    
    
End Sub

