Attribute VB_Name = "Module1"
Sub Button1_Click()
    Dim loan As Double
    Dim rate As Double
    Dim years As Double
    Dim payment As Double
    
    loan = Range("B1").Value
    rate

End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "0.02"
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "0.025"
    Range("B12:B13").Select
    Selection.AutoFill Destination:=Range("B12:B22"), Type:=xlFillDefault
    Range("B12:B22").Select
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "=RC[1]/12"
    Range("A12").Select
    Selection.AutoFill Destination:=Range("A12:A22"), Type:=xlFillDefault
    Range("A12:A22").Select
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=ABS(PMT(RC[-2],R9C2,R7C2))"
    Range("C12").Select
    Selection.AutoFill Destination:=Range("C12:C22"), Type:=xlFillDefault
    Range("C12:C22").Select
End Sub
