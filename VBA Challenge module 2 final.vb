Sub stock_market()


For Each ws In Worksheets   'loops through all worksheets in document. Need to ad "ws." before action code
'Declaring variables
Dim ticker As String
Dim ticker2 As String
Dim ticker3 As String
Dim ticker4 As String
'Dim open_price As Double   'this works as a double
'Dim close_price As String    'for some reason this variable has be be a string. Using a double doesn't work so I don't declare it at all and it works
Dim ticker_counter As Integer

'define initial variable values
ticker_counter = 2
open_price = ws.Cells(2, 3).Value  ' need initial condition for open price because this starts for loop on row 2
total_stock_vol = 0
local_max_per_change = 0
local_min_per_change = 0
local_max_stock_vol = 0


'Name title cells. Put at end to overwrite 1st iteration of title that appears
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "% Change"
ws.Cells(1, 12).Value = "Total Stock Vol"
ws.Cells(2, 15).Value = "greatest % increase"
ws.Cells(3, 15).Value = "greatest % decrease"
ws.Cells(4, 15).Value = "greatest total volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


'define what the lastrow of the worksheet it
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Starting for loop to run down column 1 checking for a ticker change and if so, defining open and closing prices
For i = 2 To lastrow
ticker = ws.Cells(i, 1).Value 'define ticker symbol

stock_vol = ws.Cells(i, 7).Value
total_stock_vol = stock_vol + total_stock_vol  'add all stock volumes to get the sum for each ticker symbol

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
    
    close_price = ws.Cells(i, 6).Value   'defines close price
    
    'Display the 1st 3 desired column information
    'Cells(ticker_counter, 14).Value = close_price 'I want each row in column 10 to be the next company symbol (ticker).
    ws.Cells(ticker_counter, 9).Value = ticker
    ws.Cells(ticker_counter, 10).Value = (close_price - open_price)
    per_change = (close_price - open_price) / open_price * 100
    ws.Cells(ticker_counter, 11).Value = per_change
    
    
    
    'Conditional statements to change cell colors green or red
        If ws.Cells(ticker_counter, 10).Value > 0 Then
        ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4 'green
        
        ElseIf ws.Cells(ticker_counter, 10).Value < 0 Then
        ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3 'red
        
        End If
        
        'Finding max and min % change and max stock vol. Use local max/min variable to store largest/smallest value encountered.
        If per_change > local_max_per_change Then
        local_max_per_change = per_change
        ws.Cells(2, 17).Value = local_max_per_change
        ticker2 = ws.Cells(i, 1).Value
        ws.Cells(2, 16).Value = ticker2
        End If
        
        If per_change < local_min_per_change Then
        local_min_per_change = per_change
        ws.Cells(3, 17).Value = local_min_per_change
        ticker3 = ws.Cells(i, 1).Value
        ws.Cells(3, 16).Value = ticker3
        End If
        
        If total_stock_vol > local_max_stock_vol Then
        local_max_stock_vol = total_stock_vol
        ws.Cells(4, 17).Value = local_max_stock_vol
        ticker4 = ws.Cells(i, 1).Value
        ws.Cells(4, 16).Value = ticker4
        End If
    
       
        
        
        
    
    ws.Cells(ticker_counter, 12).Value = total_stock_vol  '4th desired column
    total_stock_vol = 0 'reset total stock volume for next ticker symbol
    
    ticker_counter = ticker_counter + 1   'this counts the # of changes of company symbol

   
    open_price = ws.Cells(i + 1, 3).Value 'defines open price.This has to after ticker_counter because the open and close price needs to both be from the same ticker symbol
    'Cells(ticker_counter, 13).Value = open_price
    
    End If



Next i

Next ws

End Sub