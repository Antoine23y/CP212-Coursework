Attribute VB_Name = "PartA"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date:02/10/2026
' Program title: Assignment 2, Part A
' Description: Calculates Monthly Payments
'===========================================================+
Sub PartA()

    Dim PV As Double
    Dim x As Integer
    Dim rate As Double
    Dim years As Integer
    
    PV = InputBox("Please Enter Present Value: ")
    
    Range("B2").Value = PV
    rate = Range("B8").Value
    
    For x = 2 To 10
        Range("A" & x + 8).Value = x
        Range("B" & x + 8).Value = Application.WorksheetFunction.Pmt(rate / 12, x * 12, -PV)
        Next x
        
    For x = 2 To 6
        Range("C" & x + 8).Value = x / 100
        years = Range("D8").Value
        Range("D" & x + 8).Value = Application.WorksheetFunction.Pmt((x / 100) / 12, years * 12, -PV)
        Next x
        
    Range("B10:B18").NumberFormat = "$#, ##0.00"
    Range("D10:D18").NumberFormat = "$#, ##0.00"
    Range("C10:C18").NumberFormat = "0%"
    Range("A10:D18").HorizontalAlignment = xlCenter
    
    

    
    
End Sub
