Sub Alphatesting()
'Code for run Macros in all the sheets
Dim ws As Worksheet

For Each ws In Worksheets
        
    ws.Activate
    Debug.Print ws.Name
    
    Dim i As Integer
    Dim summary_row As Integer
    Dim Ticker As String
    Dim n As Integer
    Dim current As String

    n = Worksheets("A").UsedRange.Rows.Count
    summary_row = 2
    Range("I1").Value = "Ticker"

    For i = 2 To n

    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
    Ticker = Cells(i, 1).Value
    Range("I" & summary_row).Value = Ticker
    summary_row = summary_row + 1
    End If

    Next i
    
    Dim total_stock As Double

    n = Worksheets("A").UsedRange.Rows.Count
    summary_row = 2
    total_stock = 0
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To n
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    total_stock = total_stock + Cells(i, 7).Value
    Range("L" & summary_row).Value = total_stock
    summary_row = summary_row + 1
    total_stock = 0
    Else
    total_stock = total_stock + Cells(i, 7).Value
    End If
    
    Next i
    
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim open_value As Double
    Dim closed_value As Double
    Dim init As Integer
    
    init = 2
    n = Worksheets("A").UsedRange.Rows.Count
    summary_row = 2
    yearly_change = 0
    
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    
    For i = 2 To n
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    open_value = Cells(init, 3).Value
    closed_value = Cells(i, 6).Value
    yearly_change = closed_value - open_value
    percent_change = (closed_value / open_value) - 1
    Range("J" & summary_row).Value = yearly_change
    Range("K" & summary_row).Value = percent_change
    Range("K:K").NumberFormat = "0.00%"
    summary_row = summary_row + 1
    init = i + 1
    End If
    Next i
    
    'max_min_value
   
     Range("P1") = "Ticker"
     Range("Q1") = "Value"
     Range("O2") = "Greatest % Increase"
     Range("O3") = "Greatest % Decrease"
     Range("O4") = "Greatest % Total Value"
     Range("Q2") = "=MAX(C[-6])"
     Range("Q3") = "=MIN(C[-6])"
     Range("Q4") = "=MAX(C[-5])"
     
     For i = 2 To 91
     If Cells(i, 11) = Range("Q2") Then
     Ticker = Cells(i, 9).Value
     Range("P2").Value = Ticker
     ElseIf Cells(i, 11) = Range("Q3") Then
     Ticker = Cells(i, 9).Value
     Range("P3").Value = Ticker
     ElseIf Cells(i, 12) = Range("Q4") Then
     Ticker = Cells(i, 9).Value
     Range("P4").Value = Ticker
     End If
     Next i
     
     'Conditional format
     Dim count_format As Double
    
     count_format = Range("J:J").Rows.Count

     For i = 2 To 91
     If Cells(i, 10).Value < 0 Then
     Cells(i, 10).Interior.ColorIndex = 3
     Else
     Cells(i, 10).Interior.ColorIndex = 4
     End If
     Next i

Next

End Sub
