Attribute VB_Name = "Module1"
Sub TotalPage()


'A lot of the code requires things to be sorted just so.
MsgBox ("This script works best if the data is sorted by ticker, then by date.")

'Need a variable to keep track of which row to fill on the right side
Dim CurrentRow As Integer
'Need a variable to hold on to volume totals during each step of the loop
Dim VolumeTotal As Double
'Variables to track the yearly open and close for the stock
Dim MarketOpen As Double
Dim MarketClose As Double
'Autofind the last row
Dim lRow As Long
lRow = Cells(Rows.Count, 1).End(xlUp).Row


VolumeTotal = 0
CurrentRow = 2
'Store the first Market Open -
'future changes to Market Open will come when we move from stock to stock
'but since I'm putting that at the end of the loop, then I need to do this for the
'first stock value separately.
MarketOpen = Cells(2, 3).Value
Cells(2, 9).Value = Cells(2, 1).Value

'Put column labels in
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Yearly Change %"
Cells(1, 12).Value = "Total Volume"
'Turn the % change row into % format
Range("K2:K" & lRow).NumberFormat = "0.00%"

    'For loop to go through the sheet
    For i = 2 To lRow
        'Add the volume total up as you loop through
        VolumeTotal = VolumeTotal + Cells(i, 7).Value
        'When you reach the end of the current ticker, put all the values
        'on the right columns and reset variables
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'i is our last row on this ticker
            
            '--TickerSign
            'Snag the ticker before we leave
            Cells(CurrentRow, 9).Value = Cells(i, 1).Value
            
            '--Volume Total
            Cells(CurrentRow, 12).Value = VolumeTotal
            
            '--Yearly Change
            'Get the market close on our last row as our yearly end.  Only works if the sheet is sorted!
            MarketClose = Cells(i, 6).Value
            Cells(CurrentRow, 10).Value = MarketClose - MarketOpen
            
            '% change is change / original
            If MarketOpen > 0 Then 'I got a 0 divide error, so I have to dodge dead tickers on this
                Cells(CurrentRow, 11).Value = Cells(CurrentRow, 10).Value / MarketOpen
                'Conditionals to color the cells based on positive or negative
                If Cells(CurrentRow, 11).Value > 0 Then
                    Cells(CurrentRow, 11).Interior.ColorIndex = 4
                ElseIf Cells(CurrentRow, 11).Value < 0 Then
                    Cells(CurrentRow, 11).Interior.ColorIndex = 3
                End If
            End If
            'Reset your Volume Total for the next ticker in the loop
            VolumeTotal = 0
            'Prepare to print the ticker on the next line down
            CurrentRow = CurrentRow + 1
            'Get the yearly market open for the NEXT ticker
            MarketOpen = Cells(i + 1, 3).Value
        End If
    Next i

Call CheckTotals


End Sub

Sub CheckTotals()
'This runs the Hard difficulty section of the homework and is called at the end of TotalPage()

Dim IncTicker As String
Dim IncVal As Double
Dim DecTicker As String
Dim DecVal As Double
Dim lRow As Long
Dim VolTick As String
Dim VolVal As Double

IncVal = 0
IncTick = ""
DecVal = 0
DecTick = ""
VolVal = 0
VolTick = ""



'Check how many rows are in column 9 (ticker list)
lRow = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lRow 'starting on 2 to skip the header
    'Check the volume to see if it's higher than the current "VolVal", if it is, replace it and store the ticker
    If Cells(i, 12).Value > VolVal Then
        VolVal = Cells(i, 12).Value
        VolTick = Cells(i, 9).Value
    End If
    'Check the change to see if it's higher than our last highest change
    If Cells(i, 11).Value > IncVal Then
        IncVal = Cells(i, 11).Value
        IncTick = Cells(i, 9).Value
    End If
    'And check the change to see if it's lower than our last lowest change
    If Cells(i, 11).Value < IncVal Then
        DecVal = Cells(i, 11).Value
        DecTick = Cells(i, 9).Value
    End If

'So if values are the same, this will return the first highest/lowest value it comes to
    
Next i
'---- Print it to the side
Columns("N").ColumnWidth = 21
'Headers
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
'Greatest increase
Cells(2, 14).Value = "Greatest % Increase:"
Cells(2, 15).Value = IncTick
Cells(2, 16).Value = IncVal
Cells(2, 16).NumberFormat = "0.00%"
'Do it again for greatest decrease
Cells(3, 14).Value = "Greatest % Decrease:"
Cells(3, 15).Value = DecTick
Cells(3, 16).Value = DecVal
Cells(3, 16).NumberFormat = "0.00%"
'And volume
Cells(4, 14).Value = "Greatest Total Volume:"
Cells(4, 15).Value = VolTick
Cells(4, 16).Value = VolVal


End Sub

Sub Allsheets()
'For the Challenge section, this goes through and runs the script on each sheet.
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    Call TotalPage
Next ws


End Sub
