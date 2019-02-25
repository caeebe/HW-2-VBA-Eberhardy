Sub StockMarket()

'HW Instructions, Moderate option:
'Find the yearly change for each stock at opening of the year to closing at end of year
'Find the percent change
'report on the total volume of each stock
'report the ticker symbol

'Adjust the formatting to highlight positive changes in green and negative in red
'Also locate the stock with the greatest % increase, the greatest % decrease and Greatest total volume



'The Plan:

'1. For the yearly change we will need to read in the opening value of the first encounter of each stock
' we will need to compare that with the closing value of the last encounter of each stock
' we will start at the top of the column and proceed downwards as it is already sorted by ticker and date
' the percent change will just use the final yearly opening value and closing yearly value for each stock

'2. For the Total Stock Volume we will need to count a running total of the <vol> for each day for each stock

'3. We will read in the ticker name before proceeding to the next stock, and reset the values on the first day of
' each stock

'The Challange:
'adjust the code to run on each worksheet by using
'For each ws in Worksheets
'    do stuff
'next ws
'also adjust all cells and range callouts to ws.cells(x,y)...

'The Program:

'Declare all the variables before entering the loops:

'First we need to find how long each sheet is
Dim LastRow As Long

'variables for value at opening and closing of the year and for calculating the yearly & percent change for each stock
Dim OpeningValue As Double
Dim ClosingValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double

'variable for cumulatively adding up the total stock volume, it will reset for each stock just liek the above ones
Dim VolumeSum As Double

'variable for the stock name
Dim Ticker As String

'variable for counting the number of tickers, initialize at 1 as we will be inserting them starting at row 2
Dim TickCount As Long

'declare variables of greatest increase and decrease and volume, and to capture the associated tickers (Hard)
Dim GreatIncrease As Double
Dim GreatDecrease As Double
Dim GreatVolume As Double

Dim GIncTicker As String
Dim GDecTicker As String
Dim GVolTicker As String
    
'These variables are just to check if there are any matching greatest values so we don't miss two stocks
'that perform the same way  This ended up being unnecessary for this set of data.
Dim ExtraInc As Long
Dim ExtraDec As Long
Dim ExtraVol As Long
    
Dim ExtraGInc() As Double
Dim ExtraGIncTick() As String
Dim ExtraGDec() As Double
Dim ExtraGDecTick() As String
Dim ExtraGVol() As Double
Dim ExtraGVolTick() As String
    
'This will start us to loop through each sheet
Dim ws As Worksheet

For Each ws In Worksheets

    'initialize all the variables when starting each sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Ticker = "Not a stock"
    TickCount = 1
    
    'Label the new columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Now start working our way down the sheet one row at a time
    'For each row we will change the name of our Ticker if needed, add to our VolumeSum or reset the VolumeSum,
    'and either read in the opening value, the closing value or no value.

    For i = 2 To LastRow

        If Ticker <> ws.Cells(i, 1).Value Then
            'this is where we reset the values on the first day of each stock ticker name
            Ticker = ws.Cells(i, 1).Value
            OpeningValue = ws.Cells(i, 3).Value
            VolumeSum = ws.Cells(i, 7).Value
        Else
            'now we are into the middle of the year for each stock, adding to the volume and looking for the end
            VolumeSum = VolumeSum + ws.Cells(i, 7).Value
        
            'Here is where we check if the next stock is different or not
            'if the next stock is different its the last day and we print out all our values we've read in and summed
            If Ticker <> ws.Cells(i + 1, 1).Value Then
                'finally read in the closing value
                ClosingValue = ws.Cells(i, 6).Value
                
                'Calculate the yearly change now that we have both
                YearlyChange = ClosingValue - OpeningValue
                If OpeningValue <> 0 Then
                    PercentChange = YearlyChange / OpeningValue
                Else
                    PercentChange = 0
                End If
                                
                'add to the ticker counter to place the new values into the new table starting with row 2
                TickCount = TickCount + 1
                ws.Cells(TickCount, 9).Value = Ticker
                ws.Cells(TickCount, 10).Value = YearlyChange
                ws.Cells(TickCount, 11).Value = Format(PercentChange, "Percent")
                ws.Cells(TickCount, 12).Value = VolumeSum
            
                'Now do some formatting before moving on, Green if 0 or positive and Red if Negative
                If YearlyChange >= 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                Else
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                End If
            
            End If
        
        End If
        
    Next i


    'HW Instructions: Hard Option
    'Find in the new Table the Greatest %Increase, %Decrease and Volume


    'And now search by row for the biggest and smallest number
    'by comparing sizes and only putting the highest or lowest number into the variable

    'Initialize with the first value in the column to start the comparisons
    GreatIncrease = CDbl(ws.Range("k2").Value)
    GIncTicker = ws.Range("I2").Value

    GreatDecrease = CDbl(ws.Range("k2").Value)
    GDecTicker = ws.Range("I2").Value

    GreatVolume = ws.Range("L2").Value
    GVolTicker = ws.Range("I2").Value

    'Now work our way down each ticker evaluating for higher or lower values for each
    For i = 3 To TickCount
        If CDbl(ws.Cells(i, 11).Value) > GreatIncrease Then
            GreatIncrease = CDbl(ws.Cells(i, 11).Value)
            GIncTicker = ws.Cells(i, 9).Value
        
        End If
    
        If CDbl(ws.Cells(i, 11).Value) < GreatDecrease Then
            GreatDecrease = CDbl(ws.Cells(i, 11).Value)
            GDecTicker = ws.Cells(i, 9).Value

            
        End If
    
        If ws.Cells(i, 12).Value > GreatVolume Then
        GreatVolume = ws.Cells(i, 12).Value
            GVolTicker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    'use this variable to see if there are any equally greatest increases or decreases (this was unnecessary)
    ExtraTickInc = 0
    ExtraTickDec = 0
    ExtraTickVol = 0

    'Now double check that none of the tickers have matching greatest values (this was unnecessary for this data)
        For i = 3 To TickCount
        If CDbl(ws.Cells(i, 11).Value) = GreatIncrease And ws.Cells(i, 9) <> GIncTicker Then
            ExtraInc = ExtraInc + 1
            ExtraGInc(ExtraInc) = CDbl(ws.Cells(i, 11).Value)
            ExtraGIncTick(ExtraInc) = ws.Cells(i, 9).Value
        
        End If
    
        If CDbl(ws.Cells(i, 11).Value) = GreatDecrease And ws.Cells(i, 9) <> GDecTicker Then
            ExtraDec = ExtraDec + 1
            ExtraGDec(ExtraDec) = CDbl(ws.Cells(i, 11).Value)
            ExtraGDecTick(ExtraDec) = ws.Cells(i, 9).Value
        End If
    
        If ws.Cells(i, 12).Value = GreatVolume And ws.Cells(i, 9) <> GVolTicker Then
            ExtraVol = ExtraVol + 1
            ExtraGVol(ExtraVol) = CDbl(ws.Cells(i, 11).Value)
            ExtraGVolTick(ExtraVol) = ws.Cells(i, 9).Value
        End If
        
    Next i

    If ExtraInc > 0 Or ExtraDec > 0 Or ExtraVol > 0 Then
        MsgBox ("there are matching greatest values")
        'I never got a message box so this whole section was not needed for this dataset
        'If it had been necessary I would have displayed the matching Greatest data in new columns below
    End If


    'Display the results in a 3rd new table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("P2").Value = GIncTicker
    ws.Range("Q2").Value = Format(GreatIncrease, "Percent")

    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = GDecTicker
    ws.Range("Q3").Value = Format(GreatDecrease, "Percent")

    ws.Range("o4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = GVolTicker
    ws.Range("Q4").Value = GreatVolume
    
    
    'and format the columns to autofit the new tables to make it readable
    For i = 1 To 17
 
         ws.Columns(i).EntireColumn.AutoFit
 
    Next i
    
Next ws

End Sub