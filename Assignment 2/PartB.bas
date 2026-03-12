Attribute VB_Name = "PartB"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date:02/10/2026
' Program title: Assignment 2, Part B
' Description: Double Letter Words
'===========================================================+
Sub PartB()

    Dim wsWords As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim x As Long
    Dim outputRow As Long
    Dim totalWords As Long
    Dim doubleCount As Long
    Dim currentWord As String
    Dim y As Integer
    
    Set wsWords = Worksheets("Words")
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("PartB").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsNew = Worksheets.Add(After:=wsWords)
    wsNew.Name = "PartB"
    
    wsWords.Columns("A").Copy wsNew.Columns("E")
    
    lastRow = wsNew.Cells(wsNew.Rows.Count, "E").End(xlUp).Row
    
    outputRow = 2
    
    For x = 1 To lastRow
    If IsNumeric(wsNew.Cells(x, 5).Value) Then
        wsNew.Cells(outputRow, 1).Value = wsNew.Cells(x, 5).Value 'Number
    Else
        wsNew.Cells(outputRow, 2).Value = wsNew.Cells(x, 5).Value 'Word
        totalWords = totalWords + 1
        outputRow = outputRow + 1
    End If
Next x
    
    For x = 2 To outputRow - 1
        currentWord = wsNew.Cells(x, 2).Value
        
        For y = 1 To Len(currentWord) - 1
            If Mid(currentWord, y, 1) = Mid(currentWord, y + 1, 1) Then
                wsNew.Cells(x, 3).Value = currentWord
                doubleCount = doubleCount + 1
                Exit For
            End If
        Next y
    Next x
    
    wsNew.Range("B1").Value = "Words"
    wsNew.Range("C1").Value = "Double-Letter"
    With wsNew.Range("B1:C1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Name = "Times New Roman"
    End With
    
    wsNew.Columns("A:C").HorizontalAlignment = xlCenter
    wsNew.Columns("A:C").AutoFit
    wsNew.Columns("E").ClearContents
    
    MsgBox "Total Words: " & totalWords & vbNewLine & "Double-Letter Words: " & doubleCount
    

End Sub
