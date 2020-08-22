Attribute VB_Name = "Module2"
Sub SummaryTable()
Attribute SummaryTable.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ws As Worksheet

For Each ws In Worksheets 'For Each Function to Repeat on Each Worksheet
ws.Activate

    Range("U2").Select
    ActiveCell.FormulaR1C1 = "Max"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "Min"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "Absolute Variation"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Total Volume"
    Columns("V:W").Select
    Columns("V:W").EntireColumn.AutoFit
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[-5]:R[3167]C[-5])"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[-1]C[-5]:R[3166]C[-5])"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[-8]:R[3167]C[-8])"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[-1]C[-8]:R[3166]C[-8])"
    Range("W4").Select
    Columns("W:W").EntireColumn.AutoFit
    
    
Next ws


End Sub
