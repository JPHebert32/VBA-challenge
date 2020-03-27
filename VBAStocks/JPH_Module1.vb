Sub VBA_Stocks():

'define everything
Dim ws As Worksheet
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_change As Double
Dim Percent_Change As Double

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Columns("I").ColumnWidth = 6
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 16
    ws.Columns("M").ColumnWidth = 6
    ws.Columns("N").ColumnWidth = 18.5
    ws.Columns("O").ColumnWidth = 6
    ws.Columns("P").ColumnWidth = 14.5

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"



    ' Set an initial variable for holding the Ticker Symbol
    Dim Ticker As String
    ' Set an initial variable for integer
    Dim Volume As Double
    Dim Ticker_Count As Integer
    Volume = 0

    Ticker_Count = 0
    ' Keep track of the location for summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    'find Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Year_Open = ws.Cells(2, 3).Value
  
  
    ' Loop through all Tickers and Trade Date
    For i = 2 To Last_Row
    
      ' Check if we are still within the Ticker Symbol, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Set the Ticker Symbol
        Ticker = ws.Cells(i, 1).Value
      
        'Add to ticker_count
        Ticker_Count = Ticker_Count + 1
      
        ' Add to the Volume Total
        Volume = Volume + ws.Cells(i, 7).Value

        ' Print the Ticker Symbol in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
      
        ' Print the Yearly Change in the Summary Table
        Yearly_change = Year_Open - Year_Close
        
        ' Check if the Yearly_Change is greater than or equal 0
        ws.Range("J" & Summary_Table_Row).Value = Yearly_change
        If (Yearly_change >= 0) Then

        ' Color the Yearly_Change cell to green
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

        ' Check if the Yearly_change is less than or equal to 0...
        ElseIf (Yearly_change <= 0) Then

        ' Color the Yearly_Change cell to Red
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
              
        'Print the Percent Change in the Summary Table
        If (Year_Open = 0) And (Year_Close = 0) Then
            Percent_Change = 0
        Else:
            Percent_Change = (Year_Open - Year_Close) / Year_Close
            
        End If
         ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, ["0.00"] & "%")
        
        ' Print the Volume to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Volume

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the Interger Variable
        Volume = 0
        Ticker_Count = 1
        Year_Open = ws.Cells(i + Ticker_Count, 3).Value
        Year_Close = 0
      

      ' If the cell immediately following a row is the same Ticker Symbol
      Else

        ' Add to the Volume total Total
        Volume = Volume + ws.Cells(i, 7).Value
        Year_Close = ws.Cells(i + 1, 6).Value

      End If
    Next i
    
    'Challenge 1
    'Pecent Change Variables
    Dim Percent_Range As Range
    Dim Max_Increase As Double
    Dim Max_Decrease As Double


    'Max Total Volume Variables
    Dim Volume_Range As Range
    Dim Max_Volume As Double

    Dim Find_Range As Range
    Dim j, LastRow2 As Integer
    
    'Find and Assign Percent Change Max Increase
    Set Percent_Range = ws.Columns("K")
    Max_Increase = ws.Application.Max(Percent_Range)
    ws.Cells(2, 16).Value = Max_Increase
    
    'Find and Assign Percent Change Max Decrease
    Max_Decrease = ws.Application.Min(Percent_Range)
    ws.Cells(3, 16).Value = Max_Decrease
    
    'Find and Assign Max Total Volume
    Set Volume_Range = ws.Columns("L")
    Max_Volume = ws.Application.Max(Volume_Range)
    ws.Cells(4, 16).Value = Max_Volume
   
   'Challenge Tickers
    Dim Ticker_Increase As String
    Dim Ticker_Decrease As String
    Dim Ticker_Max_Vol As String
    
    LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For j = 2 To LastRow2
        If ws.Cells(j, 11) = Max_Increase Then
            Ticker_Increase = ws.Cells(j, 9).Value
        End If
        
    
        If ws.Cells(j, 11) = Max_Decrease Then
            Ticker_Decrease = ws.Cells(j, 9).Value
            
        End If
            
        If ws.Cells(j, 12) = Max_Volume Then
            Ticker_Max_Vol = ws.Cells(j, 9).Value
            
        End If
            
    Next j
    
    ws.Cells(2, 15).Value = Ticker_Increase
    ws.Cells(3, 15).Value = Ticker_Decrease
    ws.Cells(4, 15).Value = Ticker_Max_Vol

  Next ws
End Sub
