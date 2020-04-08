Attribute VB_Name = "Module1"
'Instructions
'------------------
'Create a script that will loop through one year of stock data for each run and return the total volume each stock had over
'that year.
'You will also need to display the ticker symbol to coincide with the total stock volume.

'Steps
'------------------
'Create new totals WS
'For loop to cycle data WS's
'Need to sum volume in each row and return total to totals WS, then move to next

'Code
'------------------
Sub WallStreetTicker_Embedded()
    
    'Create new totals WS
    ' Add a sheet named "StockTotals"
    'Don't want to add a new sheet...Sheets.Add.Name = "StockTotals"
    
    ' Specify the location of the stock_totals_sheet
    'Set stock_totals_sheet = Worksheets("StockTotals")
    
    'Initialize row counter variable
    Dim i As Long
    i = 0
    'Initialize last row variable
    Dim lastRow As Long
    lastRow = 0
    'Initialize totals ticker row counter
    Dim newRow As Long
    newRow = 0
    'Initialize variable to Store previous value
    Dim prevValue As Double
    
     ' Set a variable for specifying the column of ticker name
     '...Note this Hardcodes this location, may want to change later
    Dim tickername_column As Integer
    tickername_column = 1
    
    ' Set a variable for specifying the column of volume
    '...Note this Hardcodes this location, may want to change later
    Dim volume_column As Integer
    volume_column = 7
    
    ' Set a variable for specifying the column of Totals Ticker
    '...Note this Hardcodes this location, may want to change later
    Dim totalsticker_column As Integer
    totalsticker_column = 9
    
    ' Set a variable for specifying the column of Totals Volume
    '...Note this Hardcodes this location, may want to change later
    Dim totalsvolume_column As Integer
    totalsvolume_column = 10
    
  
    
    'For loop to cycle data WS's
    For Each ws In Worksheets
        'Initialize counters and new labels for Totals columns
        ws.Cells(1, totalsticker_column).Value = "Ticker"
        ws.Cells(1, totalsvolume_column).Value = "Total Volume"
        
        'Start position for newRow
        newRow = 2
        
        'Start prevValue at zero
        prevValue = 0
        
        'Need to sum volume in each row and return total to totals WS, then move to next
        'First have to get ticker name in column 1 and store in Totals Volume
       
       'Find last row
       lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
       'MsgBox ("lastRow is " & lastRow)
        ' Loop through rows in the column, skipping header row
        i = 2
        
        'Change for to lastRow when ready to run entire data...
        '!!! Crashes here
        For i = 2 To lastRow
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, tickername_column).Value <> ws.Cells(i, tickername_column).Value Then
                'Copy new ticker name to newRow
                ws.Cells(newRow, totalsticker_column).Value = ws.Cells(i, tickername_column).Value
                
                'Print totalization to cell
                ws.Cells(newRow, totalsvolume_column).Value = prevValue
                
                'Get ready for new ticker
                newRow = newRow + 1
                prevValue = 0
                
            ElseIf ws.Cells(i + 1, tickername_column).Value = ws.Cells(i, tickername_column).Value Then
                'Sum current volume total with this cell
                prevValue = ws.Cells(i, volume_column).Value + prevValue
                
            Else
                MsgBox ("Something went wrong")
            End If
        Next i 'end of loop through rows
    Next ws 'end of loop through worksheets
    
    'move created sheet to be first sheet
    'Sheets("StockTotals").Move Before:=Sheets(1)
    
    
End Sub






