Attribute VB_Name = "Module1"
Sub stockexercise()

Dim ws As Worksheet

For Each ws In Worksheets 'For Each Function to Repeat on Each Worksheet
ws.Activate
'Declaring all the variables needed for the table, including the counters that we'll need for each variable

Dim ticker As String
Dim year As Date
Dim initial As Double
initial = 0
Dim high As Double
high = 0
Dim low As Double
low = 0
Dim final As Double
final = 0
Dim total_volume As Double
total_volume = 0
Dim summary As Integer
summary = 2
Dim MaxPerc As Double
MaxPerc = 0

LastRow = ws.Cells(1, 1).End(xlDown).Row ' created a variablefor the last row to simplify the for loops later in the process

Range("J1").Value = "Ticker" ' assigned a location for each of the variables that we want to report on
Range("K1").Value = "Total Initial"
Range("L1").Value = "Total High"
Range("M1").Value = "Total Low"
Range("N1").Value = "Total Close"
Range("O1").Value = "Total Volume"
Range("Q1").Value = "Absolute Variation"
Range("R1").Value = "Percentage Variation"

Dim i As Long

For i = 2 To LastRow ' For loop to roll up the stock type and sum the values of each variable

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then '
ticker = Cells(i, 1).Value
initial = initial + Cells(i, 3).Value ' I wasn't sure if for the initial and final values we wanted to add the values or find an average. I assumed it ws the former, so just added the values. If not I would have found the way to count the number of rows and divide the values
high = high + Cells(i, 4).Value
low = low + Cells(i, 5).Value
final = final + Cells(i, 6).Value
total_volume = total_volume + Cells(i, 7).Value

Range("J" & summary).Value = ticker
Range("K" & summary).Value = initial
Range("L" & summary).Value = high
Range("M" & summary).Value = low
Range("N" & summary).Value = final
Range("O" & summary).Value = total_volume


summary = summary + 1
ticker = 0
initial = 0
high = 0
low = 0
final = 0
total_volume = 0

Else

initial = initial + Cells(i, 3).Value
high = high + Cells(i, 4).Value
low = low + Cells(i, 5).Value
final = final + Cells(i, 6).Value
total_volume = total_volume + Cells(i, 7).Value

End If

Next i

LastRow2 = ws.Cells(1, 10).End(xlDown).Row ' created a second variable for the last lrow of the new smaller table for the aggregated vaues

For i = 2 To LastRow2 ' New for loop for the different calculations

Cells(i, 17).Value = Cells(i, 14).Value - Cells(i, 11).Value

If Cells(i, 11).Value > 0 Then ' conditional to make sure we avoid divisionby zero, whih cause a runtime error

Cells(i, 17).Value = Cells(i, 14).Value - Cells(i, 11).Value
Cells(i, 18).Value = Cells(i, 17).Value / Cells(i, 11).Value
ws.Cells(i, 18).NumberFormat = "0.00%" ' Formatting percentages

Else

Cells(i, 18).Value = 0

End If

If Cells(i, 17) >= 0 Then ' coloring for the indicators
Cells(i, 17).Interior.ColorIndex = 4
Cells(i, 18).Interior.ColorIndex = 4

Else

Cells(i, 17).Interior.ColorIndex = 3
Cells(i, 18).Interior.ColorIndex = 3

End If
Next i


Next ws

End Sub

' Couldn't find the right way to find the Max and Min values (especially find the best way to set up a nested for loop to also get the stock type)

