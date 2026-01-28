' ==== CP212 Windows Application Programming ===============+
' Name:Antoine Youssef
' Student ID: 169069832
' Date: 2026-01-28
' Program title: Writing/Formatting a Spreadsheet Range
' Description:
'===========================================================+
Sub Task1()

    Range("A1").Name = "Title"
    Range("A3:F3").Name = "Headings"
    Range("A4: A21").Name = "EmpNumbers"
    Range("B4:F21").Name = "Scores"
    
    Range("Title").Font.Bold = True
    Range("Title").Font.Size = 14
    
    With Range("Headings")
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlRight
    End With
    
    Range("EmpNumbers").Font.Color = RGB(0, 0, 255)
    Range("Scores").Interior.Color = RGB(200, 200, 200)
    
    Range("A22").Value = "Averages"
    Range("A22").Font.Bold = True
    
    Range("B22").FormulaLocal = "=AVERAGE(B4:B21)"
    Range("B22").Copy
    Range("C22:F22").PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False
End Sub
