Attribute VB_Name = "Task1"
' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-01-22
' Program title: Exam Score Statstics
' Description: Reports the average, standard deviation, minimum, and maximum of the scores in a message box
'===========================================================+

Sub ExamStatistics()

    Dim avgScore As Double
    Dim stdDev As Double
    Dim minScore As Double
    Dim maxScore As Double
    Dim outputMsg As String
    
    avgScore = WorksheetFunction.Average(Range("A1:A100"))
    stdDev = WorksheetFunction.StDev(Range("A1:A100"))
    minScore = WorksheetFunction.Min(Range("A1:A100"))
    maxScore = WorksheetFunction.Max(Range("A1:A100"))
    
    outputMsg = "Exam Statistics" & vbNewLine & vbNewLine & _
                "Average: " & Round(avgScore, 2) & vbNewLine & _
                "Standard Deviation: " & Round(stdDev, 2) & vbNewLine & _
                "Minimum: " & Round(minScore, 2) & vbNewLine & _
                "Maximum: " & Round(maxScore, 2)
    
    MsgBox outputMsg, vbOKOnly, "Exam Results"
    
End Sub
