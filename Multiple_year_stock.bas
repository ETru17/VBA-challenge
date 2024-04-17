Attribute VB_Name = "Module1"
Sub mulitple_year_stock()
    
    Dim ws As Worksheet
    Dim I As Long
    Dim Ticker_symbol As String
    Dim LastRow As Long
    Dim Combined_sheet As Worksheet
    Dim tickerRow As Long
    Dim Yearly_Change As Double
    Dim Percent_change As Double
    Dim Total_Stock_Volume As Long
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    Dim GreatIncrTicker As String
    Dim GreatDecrTicker As String
    Dim GreatVolTicker As String
    
    'WorksheetName = ws.Name
    
    Sheets.Add.Name = "Combined_Data"
    Sheets("Combined_Data").Move Before:=Sheets(1)
    Set Combined_sheet = Worksheets("Combined_Data")
    
    'Create columns in first sheet
    With Combined_sheet
    
            Combined_sheet.Range("A1").Value = "ticker_symbol"
            Combined_sheet.Range("B1").Value = "Yearly Change"
            Combined_sheet.Range("C1").Value = "Percent Change"
            Combined_sheet.Range("D1").Value = "Total Stock Volume"
            
            Combined_sheet.Cells(2, 6).Value = "Greatest % Increase"
            Combined_sheet.Cells(3, 6).Value = "Greatest % Decrease"
            Combined_sheet.Cells(4, 6).Value = "Greatest Total Volume"
            Combined_sheet.Range("G1").Value = "Ticker"
            Combined_sheet.Range("H1").Value = "Value"
            
            Combined_sheet.Columns("A:H").AutoFit
        
            
     End With
    
    tickerRow = 2
     
     j = 2
     
    GreatIncr = 0
    GreatDecr = 1
    GreatVol = 0
    
    'loop through column A on every sheet to extract name of ticket
    
    For Each ws In Worksheets
      
      If ws.Name <> Combined_sheet.Name Then
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            For I = 2 To LastRow
                    
                    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                        Combined_sheet.Cells(tickerRow, 1).Value = ws.Cells(I, 1).Value
                    
                    'Yearly Change
                        Combined_sheet.Cells(tickerRow, 2).Value = ws.Cells(I, 6).Value - ws.Cells(j, 3).Value
                    
                    'Percentage Change
                    If ws.Cells(j, 3).Value <> 0 Then
                            Percent_change = (ws.Cells(I, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(I, 3)
                            Combined_sheet.Cells(tickerRow, 3).Value = Percent_change
                            Combined_sheet.Columns("C:C").NumberFormat = "0.00%"
                            
                            If Combined_sheet.Cells(tickerRow, 3).Value < 0 Then
                
                                Combined_sheet.Cells(tickerRow, 3).Interior.ColorIndex = 3
                
                             Else
                
                                Combined_sheet.Cells(tickerRow, 3).Interior.ColorIndex = 4
                            End If
                    End If
                    
                    'Total Stock Volume
                    Combined_sheet.Cells(tickerRow, 4).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(I, 7)))
                            
                    tickerRow = tickerRow + 1
                    j = I + 1
                        
                 End If
     
            Next I
       End If

Next ws

            
           
'Summary Chart Create
LastRowh = Combined_sheet.Cells(Combined_sheet.Rows.Count, 4).End(xlUp).Row
                    
For h = 2 To LastRowh
        If Combined_sheet.Cells(h, 3).Value > GreatIncr Then
            GreatIncr = Combined_sheet.Cells(h, 3).Value
            GreatIncrTicker = Combined_sheet.Cells(h, 1).Value
        End If
        
        If Combined_sheet.Cells(h, 3).Value < GreatDecr Then
            GreatDecr = Combined_sheet.Cells(h, 3).Value
            GreatDecrTicker = Combined_sheet.Cells(h, 1).Value
        End If
        
        If Combined_sheet.Cells(h, 4).Value > GreatVol Then
            GreatVol = Combined_sheet.Cells(h, 4).Value
            GreatVolTicker = Combined_sheet.Cells(h, 1).Value
        End If
 Next h
                                
Combined_sheet.Cells(4, 7).Value = GreatVolTicker
Combined_sheet.Cells(4, 8).Value = GreatVol
Combined_sheet.Cells(2, 7).Value = GreatIncrTicker
Combined_sheet.Cells(2, 8).Value = GreatIncr
Combined_sheet.Cells(3, 7).Value = GreatDecrTicker
Combined_sheet.Cells(3, 8).Value = GreatDecr

  ' Autofit columns
Combined_sheet.Columns("A:H").AutoFit
                            
    
        

End Sub











'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock. The result should match the following image:


