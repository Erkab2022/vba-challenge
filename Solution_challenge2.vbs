Attribute VB_Name = "Module1"
Sub Yearlychange():
'-----------------------------
'Loop for each Worksheet
'----------------------------
For Each ws In Worksheets

' Determine the Last Row
        Lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set a variable for the tickername
    Dim tickername As String

    'Set an initial variable for yearly change
    Dim yearly_change As Double
    yearly_change = 0
    
    'Set an initial variable for the open
    Dim open_price As Double
    open_price = ws.Cells(2, 3).Value

    
    'Set an initial variable for the total volume
    Dim total_volume As LongPtr
    total_volume = 0
    
    'Set an initial variable for the percentage_change
    Dim percentage_change As Double
    percentage_change = 0

    'Track each location of the tickername in the summary table
    Dim summary_table_row As LongPtr
        summary_table_row = 2
    
    'Set the differents titles of the summary table
    ws.Range("L1").Value = "Ticker"
    ws.Range("M1").Value = "Yearly Change"
    ws.Range("N1").Value = "Percentage change"
    ws.Range("O1").Value = "Total Stock Volume"

    'Loop for all ticker
    Dim i As LongPtr
        For i = 2 To Lastrow1
    
    'Check if it still the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Set the tickername
    tickername = ws.Cells(i, 1).Value

    'Add to total yearly_change
    yearly_change = yearly_change + ws.Cells(i, 6).Value - open_price
    
    'Add the total_volume
    total_volume = total_volume + ws.Cells(i, 7).Value

    'Add the total pecentage_change
    percentage_change = yearly_change / open_price
    
    'Print the tickername in the summary table
    ws.Range("L" & summary_table_row).Value = tickername

    'Print the yearly-change in the summary table
    ws.Range("M" & summary_table_row).Value = yearly_change
    
    'Print the Percentage-change in the summary table
    ws.Range("N" & summary_table_row).Value = percentage_change
    
    'change the format of the Range N
    ws.Range("N" & summary_table_row).NumberFormat = "0.00%"
    
    'Print the total volume in the summary table
    ws.Range("O" & summary_table_row).Value = total_volume

    'Add one to the summary table
    summary_table_row = summary_table_row + 1

    'Reset the yearly_change
    yearly_change = 0
    
    'Reset the total volume
    total_volume = 0
    
    'Reset the percentage_change
    percentage_change = 0
    
    'Reset the open_price
    open_price = ws.Cells(i + 1, 3).Value
    
    'If the cell immediately following is still the same
    Else
    
    On Error Resume Next
    
    'Add to the total_volume
    total_volume = (total_volume) + ws.Cells(i, 7).Value
    
   End If
    
    
Next i

'--------------------------------
'Insert color to the percentage
'-----------------------------------
'Detrrmine the Last Row
    Lastrow2 = ws.Cells(Rows.Count, 14).End(xlUp).Row

'Define a variable
For j = 2 To Lastrow2

'Make the condition
If ws.Cells(j, 13).Value > 0 Then

'Fill the cell with the green color
ws.Cells(j, 13).Interior.ColorIndex = 4
'Otherwise
Else
'Fill the cell with the red color
ws.Cells(j, 13).Interior.ColorIndex = 3
End If
Next j

'---------------------------------
'Retrieve the Greatest value
'-------------------------------

ws.Range("S2") = WorksheetFunction.Max(ws.Range("N2:N" & Lastrow2))
ws.Range("S3") = WorksheetFunction.Min(ws.Range("N2:N" & Lastrow2))
ws.Range("S4") = WorksheetFunction.Max(ws.Range("O2:O" & Lastrow2))

'Set the titles of the range
ws.Range("R1") = "Ticker"
ws.Range("S1") = "Value"

'Loop to find the ticker
Dim n As Integer
For n = 2 To Lastrow2

'Make the condition to retrieve the greatest% Increase
If ws.Cells(n, 14).Value = ws.Range("S2") Then

'Find the value and give the title
ws.Range("Q2").Value = "Greatest% Increase"
ws.Range("R2") = ws.Cells(n, 12).Value
End If

'Make the condition to retrieve the greatest% decrease
If ws.Cells(n, 14).Value = ws.Range("S3") Then
ws.Range("Q3").Value = "Greatest% Decrease"
ws.Range("R3").Value = ws.Cells(n, 12).Value


'Convert the format to percentage format
ws.Range("S2:S3").NumberFormat = "0.00%"
End If

'Make the condition to retrieve the greatest total volume
If ws.Cells(n, 15).Value = ws.Range("S4") Then
ws.Range("Q4").Value = "Greatest Total volume"
ws.Range("R4").Value = ws.Cells(n, 12).Value
End If

Next n

Next ws


End Sub
