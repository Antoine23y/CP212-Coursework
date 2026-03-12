VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormGasPrices 
   Caption         =   "Gas Prices Over Last 20 Years"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5985
   OleObjectBlob   =   "FormGasPrices.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormGasPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-03-12
' Program title: Assignment 4
'===========================================================+
Private Sub UserForm_Initialize()
    Cities.AddItem "Ottawa"
    Cities.AddItem "Toronto West"
    Cities.AddItem "Toronto East"
    Cities.AddItem "Windsor"
    Cities.AddItem "London"
    Cities.AddItem "Sudbury"
    Cities.AddItem "Sault Saint Marie"
    Cities.AddItem "Thunder Bay"
    Cities.AddItem "North Bay"
    Cities.AddItem "Timmins"
    
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()

    Dim x As Integer
    Dim cityChosen As Boolean
    
    cityChosen = False
    
    For x = 0 To Cities.ListCount - 1
        If Cities.Selected(x) = True Then
            cityChosen = True
        End If
    Next x
    
    If cityChosen = False Then
        MsgBox "Please Select at least one City!"
        Exit Sub
    End If
    
    If chkHighest.Value = False And chkLowest.Value = False Then
        MsgBox "Please select at least one Measure!"
        Exit Sub
    End If
    
    Call ShowResults
    
    If optYes.Value = True Then
        Call MakeChart
    End If
    
End Sub
