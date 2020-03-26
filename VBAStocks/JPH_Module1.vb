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
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

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

        ' Color the Yearly_Change grade green
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

        ' Check if the student's grade is less than or equal to 0...
        ElseIf (Yearly_change <= 0) Then

        ' Color the Yearly_Change Red
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
              
        'Print the Percent Change in the Summary Table
        If (Year_Open = 0) And (Year_Close = 0) Then
            Percent_Change = 0
        Else:
            Percent_Change = (Year_Open - Year_Close) / Year_Close
            
        End If
         ws.Range("K" & Summary_Table_Row).Value = Percent_Change

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
  Next ws
End Sub