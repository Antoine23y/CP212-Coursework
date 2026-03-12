Attribute VB_Name = "Task1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 1690969832
' Date: 2026-02-27
' Program title: Assignment 3, Task 1
' Description:
'===========================================================+
Sub OneSubProgram()

    Application.Volatile
    Randomize
    
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim n As Double
    Dim BW As Double
    Dim x As Long, y As Long
    Dim randNum As Double
    Dim index As Integer
    
    a = Range("B2").Value
    b = Range("B3").Value
    c = Range("B4").Value
    n = Range("B5").Value
    
    BW = (c - a) / 20
    
    Range("A8:D10000").ClearContents
    
    Range("A8").Value = "Random Number"
    Range("B8").Value = "Index"
    Range("C8").Value = "Bin"
    Range("D8").Value = "Frequency"
    
    Dim freq(1 To 20) As Long
    
    For x = 1 To n
    
        randNum = Triangular(a, b, c)
        
        Range("A8").Offset(x, 0).Value = randNum
        
        index = Int((randNum - a) / BW) + 1
        
        If index < 1 Then index = 1
        If index > 20 Then index = 20
        
        freq(index) = freq(index) + 1
        
    Next x
    
    For y = 1 To 20
        Range("B8").Offset(y, 0).Value = y
        
        Range("C8").Offset(y, 0).Value = "[" & _
            a + BW * (y - 1) & "-" & a + BW * y & "]"
            
        Range("D8").Offset(y, 0).Value = freq(y)
    Next y
End Sub

Function Triangular(a As Double, b As Double, c As Double) As Double

Dim U As Double
Dim d As Double

U = Rnd()
d = (b - a) / (c - a)

If U < d Then
    Triangular = a + (c - a) * Sqr(d * U)
Else
    Triangular = a + (c - a) * (1 - Sqr((1 - d) * (1 - U)))
End If

End Function
