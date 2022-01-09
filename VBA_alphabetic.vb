Attribute VB_Name = "Module1"
Sub Stocks()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'Define veriables for given info
Dim ticker_name As String
Dim open_price As Double
Dim close_price As Double


'Define Veriables for new info
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double


    'Adding extra columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Set the beginning count for the loop
    Yearly_Change = 0
    Total_Stock_Volume = 0
    ticker_name = 0
    open_price = 0
    close_price = 0

'To go to the last row
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Creating For Loops
'For Loop for the Ticker
    For i = 2 To LastRow
        
        'Define the ticker name
        Ticker = Cells(i, 1).Value
        
        'Opening price at the beinning of the year
          If open_price = 0 Then
          open_price = Cells(i, 3).Value
        
        End If
        
        'Total stock volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        'If next ticker and current ticker is not the same
        If Cells(i + 1, 1).Value <> Ticker Then
        'To get to the next ticker name
        ticker_name = ticker_name + 1
        Cells(ticker_name + 1, 9) = Ticker
        
        'Closing price at the end of the year
        close_price = Cells(i, 6)

        'To culculate Yearly Change
        Yearly_Change = close_price - open_price
        'Add to the Yearly Change cell
        Cells(ticker_name + 1, 10).Value = Yearly_Change
        
'If change is >= 0, shade cells green
If Yearly_Change >= 0 Then
    Cells(ticker_name + 1, 10).Interior.ColorIndex = 4
'If change is < 0, shade cells red
    ElseIf Yearly_Change < 0 Then
    Cells(ticker_name + 1, 10).Interior.ColorIndex = 3
     
     End If
        
         'To find percent change value per ticker
            If open_price = 0 Then
            Percent_Change = 0
            Else: Percent_Change = (Yearly_Change / open_price)
            End If
            
            'Format Persent_Change value as %
            Cells(ticker_name + 1, 11).Value = Format(Percent_Change, "Percent")
            
            'Continue to the next ticker
            open_price = 0
            
            'Add to the Total Stock Volume
            Cells(ticker_name + 1, 12).Value = Total_Stock_Volume
            
            'Continue to the next ticker
            Total_Stock_Volume = 0
            
End If

Next i


    'Create new table for % change
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
       
 
        'Set the correct home for %  data to look through
        GPI = Cells(2, 11).Value
        GPIT = Cells(2, 9).Value
        GPD = Cells(2, 11).Value
        GPDT = Cells(2, 9).Value
        GSV = Cells(2, 12).Value
        GSVT = Cells(2, 9).Value
        
            'To go to the last row in Ticker
             LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
            
            'Run for loop
            For i = 2 To LastRow
            
            'Find Ticker with GPI
            'Turn result into % and insert into the correct cell
            
            If Cells(i, 11).Value > GPI Then
            GPI = Cells(i, 11).Value
            GPIT = Cells(i, 9).Value
            ws.Range("P2").Value = Format(GPIT, "Percent")
            ws.Range("Q2").Value = Format(GPI, "Percent")
            
            End If
            
            'Find Ticker with GPD
            'Turn result into % and insert into the correct cell
            
            If Cells(i, 11).Value < GPD Then
            GPD = Cells(i, 11).Value
            GPDT = Cells(i, 9).Value
            ws.Range("P3").Value = Format(GPDT, "Percent")
            ws.Range("Q3").Value = Format(GPD, "Percent")
            
            End If
            
          
            'Find Ticker with GSV
            'Turn result into % and insert into the correct cell
            
            If Cells(i, 12).Value > GSV Then
            GSV = Cells(i, 12).Value
            GSVT = Cells(i, 9).Value
            ws.Range("P4") = GSVT
            ws.Range("Q4") = GSV
            
           End If
        
        Next i
        
            

Next ws


End Sub
