Attribute VB_Name = "Task1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 02/10/2026
' Program title: Arrays
' Description: Finding flights
'===========================================================+

Sub findFlights()

    Dim originSearch() As String
    Dim destSearch() As String
    Dim flightsOrigin() As String
    Dim flightsDest() As String
    Dim flightsNumber() As String
    Dim numOriginSearch As Long
    Dim numDestSearch As Long
    Dim numFlights As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim count As Long

    numOriginSearch = Range("E1", Range("E1").End(xlDown)).Rows.count
    ReDim originSearch(1 To numOriginSearch)
    numDestSearch = Range("F1", Range("F1").End(xlDown)).Rows.count
    ReDim destSearch(1 To numDestSearch)

    For x = 1 To numOriginSearch
        originSearch(x) = Range("E1").Offset(x - 1, 0).Value
        Next x
    
    For y = 1 To numDestSearch
        destSearch(y) = Range("F1").Offset(y - 1, 0).Value
        Next y
    numFlights = Range("A1", Range("A1").End(xlDown)).Rows.count
    ReDim flightsOrigin(1 To numFlights)
    ReDim flightsDest(1 To numFlights)
    ReDim flightsNumber(1 To numFlights)
    
    For z = 1 To numFlights
        flightsOrigin(z) = Range("A1").Offset(z - 1, 0).Value
        flightsDest(z) = Range("B1").Offset(z - 1, 0).Value
        flightsNumber(z) = Range("C1").Offset(z - 1, 0).Value
        Next z
    
    Range("H:J").ClearContents
    Range("H1").Value = "Origin"
    Range("I1").Value = "Destination"
    Range("J1").Value = "Flights Number"
    With Range("H1:J1").Font
        .Bold = True
        .Italic = True
        .Color = vbBlue
    End With

    count = 0
    For x = 1 To numOriginSearch
        For y = 1 To numDestSearch
            For z = 1 To numFlights
                If flightsOrigin(z) = originSearch(x) And flightsDest(z) = destSearch(y) Then
                    count = count + 1
                    Range("H1").Offset(count, 0).Value = flightsOrigin(z)
                    Range("H1").Offset(count, 1).Value = flightsDest(z)
                    Range("H1").Offset(count, 2).Value = flightsNumber(z)
                End If
            Next z
        Next y
    Next x
    
    MsgBox count & " flights found!"

End Sub

