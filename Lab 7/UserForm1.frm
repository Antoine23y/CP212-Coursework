VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4995
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   Dim sizeChoice As Integer
   Dim columnChoice As Integer
   Dim x As Long
   Dim total As Double
   Dim count As Long
   Dim lastRow As Long
   Dim ws As Worksheet
   
   Set ws = ThisWorkbook.Worksheets("Data")
   
   If OptionButton1.Value Then sizeChoice = 1
   If OptionButton2.Value Then sizeChoice = 2
   If OptionButton3.Value Then sizeChoice = 3
   
   If OptionButton4.Value Then columnChoice = 2
   If OptionButton5.Value Then columnChoice = 3
   lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
   
   For x = 2 To lastRow
    If ws.Cells(x, 1).Value = sizeChoice Then
        total = total + ws.Cells(x, columnChoice).Value
        count = count + 1
    End If
   Next x
   
   If count = 0 Then
        MsgBox "No matching customers found!"
    Else
        MsgBox "Average: $" & Format(total / count, "0.00")
    End If
    
   
End Sub

Private Sub CommandButton2_Click()
    Unload Me

End Sub
