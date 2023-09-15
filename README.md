# VBA-challenge
Module 2 Homework 2 VBA Scripting Multi Year Stock Data

'begin code
Sub Stock()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Define variables
    Dim ticker As String
    Dim openvalue As Double
    Dim yearlychange As Double
    yearlychange = 0
    Dim percentchange As Double
    percentchange = 0
    Dim closevalue As Double
    closevalue = 0
    Dim volume As Double
    volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Heading Names for the Table
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
    Dim OpenTrigger As Boolean
    
        If OpenTrigger = False Then
            openvalue = Cells(i, 3).Value
            OpenTrigger = True
        End If
    
        'Ticker change and record in Summary Table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            Range("J" & Summary_Table_Row).Value = ticker
            
         'Yearly Change and record in Summary Table
            yearlychange = Cells(i, 6).Value - openvalue
            Range("K" & Summary_Table_Row).Value = yearlychange
            Range("K" & Summary_Table_Row).NumberFormat = "0.00"
            
         'Percent Change and record in Summary Table
            percentchange = (yearlychange / openvalue)
            Range("L" & Summary_Table_Row).Value = percentchange
            Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
            
         'Total Volume and record in Summary Table
            volume = volume + Cells(i, 7).Value
            Range("M" & Summary_Table_Row).Value = volume
            Range("M" & Summary_Table_Row).NumberFormat = General
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            volume = 0
            OpenTrigger = False
        Else
            volume = volume + Cells(i, 7).Value
            yearlychange = Cells(i, 6).Value - openvalue
            'percentchange = (yearlychange / openvalue) * 100

        End If

    Next i
    
    'Color Column K red if negative and green if positive
    For i = 2 To 3001
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        ElseIf Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i
    
    'Create another table to represent the greatest % increase/decrease, total volume
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    Range("R2") = WorksheetFunction.Max(Range("L2:L3001"))
    Range("R2").NumberFormat = "0.00%"
    Range("R3") = WorksheetFunction.Min(Range("L2:L3001"))
    Range("R3").NumberFormat = "0.00%"
    Range("R4") = WorksheetFunction.Max(Range("M2:M3001"))
    Range("R4").NumberFormat = "0.00E+00"
    
    'Column Q to have the Ticker populate based on Column R value
    Dim MaxP As Double
    Dim MinP As Double
    Dim MaxV As Double
    
    MaxP = Range("R2").Value
    MinP = Range("R3").Value
    MaxV = Range("R4").Value
    
    For i = 2 To 3001
        If Cells(i, 12).Value = MaxP Then
            Range("Q2").Value = Cells(i, 10).Value
        ElseIf Cells(i, 12).Value = MinP Then
            Range("Q3").Value = Cells(i, 10).Value
        ElseIf Cells(i, 13).Value = MaxV Then
            Range("Q4").Value = Cells(i, 10).Value
        End If
    Next i
        
Next ws
End Sub
'end code


'Mod2 Resources to help me:
'https://stackoverflow.com/questions/43738802/how-to-apply-vba-code-to-all-worksheets-in-the-workbook
'https://stackoverflow.com/questions/59461571/how-do-i-keep-initial-value-in-a-for-loop
