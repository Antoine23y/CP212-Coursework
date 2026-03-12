Attribute VB_Name = "GasPrices"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-03-12
' Program title: Assignment 4
'===========================================================+
Sub RunProgram()
    FormGasPrices.Show
End Sub

Sub ShowResults()
    
    Dim ws As Worksheet
    Dim x As Integer
    Dim y As Long
    Dim cityName As String
    Dim resultText As String
    Dim highestPrice As Double
    Dim lowestPrice As Double
    Dim highestDate As Variant
    Dim lowestDate As Variant
    Dim lastRow As Long
    Dim col As Integer
    Dim priceValue As Double
    
    Set ws = Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    resultText = ""
    
    For x = 0 To FormGasPrices.Cities.ListCount - 1
        If FormGasPrices.Cities.Selected(x) = True Then
        
        cityName = FormGasPrices.Cities.List(x)
        col = 0
        For y = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If ws.Cells(1, y).Value = cityName Then
                col = y
                Exit For
            End If
        Next y
        
        highestPrice = -1
        lowestPrice = 9999999
        
        For y = 2 To lastRow
            If IsNumeric(ws.Cells(y, col).Value) Then
                priceValue = ws.Cells(y, col).Value
                
                If priceValue > 0 Then
                    If priceValue > highestPrice Then
                        highestPrice = priceValue
                        highestDate = ws.Cells(y, 1).Value
                    End If
                    
                    If priceValue < lowestPrice Then
                        lowestPrice = priceValue
                        lowestDate = ws.Cells(y, 1).Value
                    End If
                End If
            End If
        Next y
        
        resultText = resultText & cityName & ":" & vbCrLf
        
        If FormGasPrices.chkHighest.Value = True Then
            resultText = resultText & "The Highest Gas Price Is $" & _
            Format(highestPrice / 100, "0.00") & " Per Liter On " & _
            Format(highestDate, "yyyy-mm-dd") & vbCrLf
        End If
        
        If FormGasPrices.chkLowest.Value = True Then
            resultText = resultText & "The Lowest Gas Price Is $" & _
            Format(lowestPrice / 100, "0.00") & " Per Liter On " & _
            Format(lowestDate, "yyyy-mm-dd") & vbCrLf
        End If
        
        resultText = resultText & vbCrLf
    End If
    Next x
    
    MsgBox resultText, vbInformation, "Gas Prices"
    
End Sub
Sub MakeChart()

    Dim wsData As Worksheet
    Dim wsGraph As Worksheet
    Dim lastRow As Long
    Dim x As Integer
    Dim y As Long
    Dim cityName As String
    Dim col As Integer
    Dim cht As ChartObject
    
    Set wsData = Worksheets("Data")
    Set wsGraph = Worksheets("MyGraph")
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    On Error Resume Next
    wsGraph.ChartObjects(1).Delete
    On Error GoTo 0
    
    Set cht = wsGraph.ChartObjects.Add(Left:=10, Width:=800, Top:=80, Height:=400)
    
    cht.Chart.ChartType = xlLine
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "Last 20 years Gas Prices"
    cht.Chart.HasLegend = True
    
    cht.Chart.ChartArea.Interior.Color = RGB(192, 192, 192)
    
    With cht.Chart.Axes(xlValue)
        .MinimumScale = 70
        .MaximumScale = 250
        .HasTitle = True
        .AxisTitle.Text = "Cents Per Liter"
    End With
    
    For x = 0 To FormGasPrices.Cities.ListCount - 1
        If FormGasPrices.Cities.Selected(x) = True Then
            cityName = FormGasPrices.Cities.List(x)
            col = 0
            
            For y = 2 To wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
                If wsData.Cells(1, y).Value = cityName Then
                    col = y
                    Exit For
                End If
            Next y
            
            If col > 0 Then
                cht.Chart.SeriesCollection.NewSeries
                With cht.Chart.SeriesCollection(cht.Chart.SeriesCollection.Count)
                    .Name = cityName
                    .XValues = wsData.Range("A2:A" & lastRow)
                    .Values = wsData.Range(wsData.Cells(2, col), wsData.Cells(lastRow, col))
                End With
            End If
        End If
    Next x
        
End Sub






