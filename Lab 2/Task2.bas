Attribute VB_Name = "Task2"
' ==== CP212 Windows Application Programming ===============+
' Name: Antoine Youssef
' Student ID: 169069832
' Date: 2026-01-22
' Program title: BMI CALCULATOR
' Description: Computes Body Mass Index (BMI) by asking for the user's height (in meters) and weight (in kg)
'===========================================================+
Sub CalculateBMI()

    Dim height As Single
    Dim weight As Single
    Dim bmi As Single
    Dim heightInput As String
    Dim weightInput As String
    
    heightInput = InputBox("Enter your height in meters (e.g., 1.85m): ", "HeightInput")
    height = CSng(heightInput)
    
    weightInput = InputBox("Enter your weight in kilograms (e.g., 75kg): ", "WeightInput")
    weight = CSng(weightInput)
    
    bmi = weight / (height ^ 2)
    
    MsgBox "Height: " & height & " meters" & vbNewLine & _
            "Weight: " & weight & " kg" & vbNewLine & _
            "BMI: " & Round(bmi, 2), _
            vbOKOnly, "BMI Results"
    



End Sub
