Attribute VB_Name = "Task2"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-03-10
' Program title: Lab 8 Task 2
' Description:
'===========================================================+
Sub OpenAFile()
    Dim strName As String
    
    On Error GoTo ErrorHandler
    
    strName = InputBox("Enter a File Full Location", "User Input")
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=strName
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "File can't be found or another error occured!"
    
End Sub

Sub ShowForm()
    UserForm1.Show
End Sub
