Attribute VB_Name = "Task1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-02-27
' Program title: Lab 6 Task 1
' Description:
'===========================================================+
Public Function NumericInText(ByVal stringValue As String) As Long

    Dim x As Long
    Dim char As String
    Dim nums As String
    
    nums = ""
    
    For x = 1 To Len(stringValue)
        char = Mid$(stringValue, x, 1)
        If IsNumeric(char) Then
             nums = nums & char
        End If
    Next x
    
    If nums = "" Then
        NumericInText = 0
    Else
        NumericInText = CLng(nums)
    End If
    
End Function
