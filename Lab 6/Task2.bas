Attribute VB_Name = "Task2"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-02-27
' Program title: Lab 6 Task 2
' Description:
'===========================================================+

Public Function ArrayEqual(ByRef a() As Single, ByRef b() As Single) As Boolean

    Dim x As Long
    
    If (LBound(a) <> LBound(b)) Or (UBound(a) <> UBound(b)) Then
        ArrayEqual = False
    End If
    
    For x = LBound(a) To UBound(a)
        If a(x) <> b(x) Then
            ArrayEqual = False
            Exit Function
        End If
    Next x
    
    ArrayEqual = True

End Function

Public Sub TestMyFunctions()
    Dim arr1(1 To 3) As Single
    Dim arr2(1 To 3) As Single
    Dim isEq As Boolean
    
    arr1(1) = 3
    arr1(2) = 5
    arr1(3) = 6
    
    arr2(1) = 4
    arr2(2) = 3
    arr2(3) = 6
    
    isEq = ArrayEqual(arr1, arr2)
    
    Worksheets(1).Range("B12").Value = isEq
    
End Sub
