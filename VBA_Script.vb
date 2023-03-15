Sub StockData()
    Dim r, rng, rngminmaxperc, rngmaxvol, FndRng   As Range
    Dim k, maxPerc, minPerc      As Variant
    Dim d           As Object
    Dim i As Integer
    Dim compareDate, lRow As Long
    Dim vol, maxvol        As LongLong
    Dim openPrice, closePrice, yearlyChange  As Double
    
    ' Loop though all the worksheets
    For Each ws In Sheets
        Set d = CreateObject("Scripting.Dictionary")
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Set rng = ws.Range("A2:A" & lastRow)
        i = 0
        
        ' Assign distinct Ticker values to the scripting object
        For Each r In rng
            If Not IsEmpty(r) Then d(r.Value) = r.Value
        Next
        
        'Assigning header values
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "YearlyChange"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        ' Loop thorugh all Ticker values
        For Each k In d.Keys
            vol = 0
            i = i + 1
            minDate = 99999999
            maxDate = 0
            ws.Cells(i + 1, 10) = d(k)        ' Populate ticker value
            For Each r In rng
                If k = r.Value Then
                    vol = vol + r.Offset(0, 6).Value        ' Loop through all the rows for one Ticker and sum up the values
                    compareDate = CLng(r.Offset(0, 1).Value)        ' Converting Date to LONG as it is stored as string in excel
                    ' Find the minimum and the maximum dates
                    If compareDate < minDate Then
                        minDate = compareDate
                    End If
                    If compareDate > maxDate Then
                        maxDate = compareDatea
                    End If
                    ' If the Date in the row = Minimum Date, assign open Price against that row
                    If compareDate = minDate Then
                        openPrice = r.Offset(0, 2).Value
                        ' ws.Cells(i + 1, 14) = r.Offset(0, 2).Value
                        ' If the Date in the wow = Maximum Date, assign close Price against that row
                    ElseIf compareDate = maxDate Then
                        closePrice = r.Offset(0, 5).Value
                        '    ws.Cells(i + 1, 15) = r.Offset(0, 2).Value
                    End If
                 End If
            Next
            yearlyChange = closePrice - openPrice
            ws.Cells(i + 1, 11) = yearlyChange        ' Populate yearly change
            ' Handle zeroes in Open and Close Price to calculate the Change % and populate it
            If openPrice = 0 And closePrice = 0 Then
                ws.Cells = 0
            ElseIf openPrice = 0 Then
                ws.Cells(i + 1, 12) = (-yearlyChange / closePrice) * 100
            Else
                ws.Cells(i + 1, 12) = (yearlyChange / openPrice) * 100
            End If
   
            ws.Cells(i + 1, 13) = vol        ' Populate volume
            ' Color the cell Red if the yearly Change is negative and green if positive
            If ws.Cells(i + 1, 11).Value > 0 Then
                ws.Range("K" & i + 1).Interior.ColorIndex = 4
            ElseIf ws.Cells(i + 1, 11).Value < 0 Then
                ws.Range("K" & i + 1).Interior.ColorIndex = 3
            End If
        Next
        lastRow = d.Count + 1
        ws.Range("L2:L" & lastRow).NumberFormat = "0.00\%"
        ws.Range("M2:M" & lastRow).NumberFormat = "0"
        ' Assign labels
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        
        ' Find out the min and max percentage change in price and corresponding ticker
        Set rngminmaxperc = ws.Range("L2:L" & lastRow)
        maxPerc = WorksheetFunction.Max(rngminmaxperc)
        minPerc = WorksheetFunction.Min(rngminmaxperc)
        Set FndRng = rngminmaxperc.Find(what:=maxPerc)
        ws.Range("Q2").Value = FndRng.Offset(0, -2).Value
        Set FndRng = rngminmaxperc.Find(what:=minPerc)
        ws.Range("Q3").Value = FndRng.Offset(0, -2).Value
        
        ' Populate the min and max percentage change in price
        ws.Range("R2").Value = maxPerc
        ws.Range("R3").Value = minPerc
        ws.Range("R2:R3").NumberFormat = "0.00\%"

        
        ' Find out the max volume and and corresponding ticker
        Set rngmaxvol = ws.Range("M2:M" & lastRow)
        maxvol = WorksheetFunction.Max(rngmaxvol)
        Set FndRng = rngmaxvol.Find(what:=maxvol)
        
        ' Populate the max volume and corresponding ticker
        ws.Range("R4").Value = maxvol
        ws.Range("R4").NumberFormat = "0"
        ws.Range("Q4").Value = FndRng.Offset(0, -3).Value
        
    Next ws
End Sub