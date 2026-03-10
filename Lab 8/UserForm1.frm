VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "User Input"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-03-10
' Program title: Lab8
' Description:
'===========================================================+

Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub txtLastName_Change()

    Dim lastName As String
    lastName = txtLastName.Text

    If Len(lastName) > 0 Then
        If Not (UCase(Right(lastName, 1)) >= "A" And UCase(Right(lastName, 1)) <= "Z") Then
            
            MsgBox "Only enter letters!", vbCritical, "Invalid"
            
            txtLastName.Text = Left(lastName, Len(lastName) - 1)
            txtLastName.SetFocus
            
            Exit Sub
        End If
    End If

End Sub
