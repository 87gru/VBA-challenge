Sub ticker()

Dim i As LongLong
Dim ws As Worksheet

    For Each ws In Worksheets '-- loop through worksheets
    
        ' Variables
        Dim TickerName As String
        Dim open_price As Double
        Dim close_price As Double
        Dim qc As Double '--quarterly change
        Dim pc As Double '--percent change
        Dim total_vol As LongLong '--total stock volume
        Dim lastrow As LongLong '-- row counter
        
        
        Dim t As Integer

        open_price = ws.Range("C2").Value
        close_price = 0
        qc = 0
        total_vol = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        t = 2
        
        '--Loop/Conditional
        
        For i = 2 To lastrow

        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then

            TickerName = ws.Cells(i, 1).Value
            total_vol = total_vol + ws.Cells(i, 7).Value
            close_price = ws.Cells(i, 6).Value
            qc = close_price - open_price
            pc = qc / open_price

            ws.Cells(t, 9).Value = TickerName
    
                If qc > 0 Then
                ws.Cells(t, 10).Value = qc
                ws.Range("J" & t).Interior.ColorIndex = 4
        
                ElseIf qc < 0 Then
                ws.Cells(t, 10).Value = qc
                ws.Range("J" & t).Interior.ColorIndex = 3

                Else
                ws.Cells(t, 10).Value = qc
                ws.Cells(t,10).NumberFormat = "0.00"

                End If
    
            ws.Cells(t, 11).Value = pc
            ws.Cells(t, 11).NumberFormat = "0.00%"
            ws.Cells(t, 12).Value = total_vol
    
            close_price = 0
            open_price = 0
            qc = 0
            pc = 0
            total_vol = 0
            t = t + 1
            open_price = ws.Cells(i + 1, 3).Value
    
        Else:
    
        total_vol = total_vol + ws.Cells(i, 7).Value
    
    End If
    
    Next i

'Name columns for summary data

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quartlery Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Autofit columns to make data more legible
ws.Columns("I:L").AutoFit


'-------------------------------------------

' Variables for Max/Min Section
Dim m As Integer
Dim sumlr As Integer ' last row counter for data in columns I - L
Dim tmax As String ' ticker for greatest % increase
Dim tmin As String ' ticker for greatest% decrease
Dim tgtv As String 'ticker for greatest total val
Dim gtv As LongLong 'max total val
Dim max As Double ' max % increase
Dim min As Double 'min $ dec
max = 0
min = 0
gtv = 0
sumlr = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'loop + conditional
    For m = 2 To sumlr

        If ws.Cells(m, 11).Value > max Then
        max = ws.Cells(m, 11).Value
        tmax = ws.Cells(m, 9).Value
        End If

        If ws.Cells(m, 11).Value < min Then
        min = ws.Cells(m, 11).Value
        tmin = ws.Cells(m, 9).Value
        End If

        If ws.Cells(m, 12).Value > gtv Then
        gtv = ws.Cells(m, 12).Value
        tgtv = ws.Cells(m, 9).Value
        End If

    Next m

'Insert values into Columns O - P
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("P2").Value = tmax
    ws.Range("Q2").Value = max
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = tmin
    ws.Range("Q3").Value = min
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = tgtv
    ws.Range("Q4").Value = gtv

ws.Columns("O:Q").AutoFit



Next ws
    
        
 

End Sub



