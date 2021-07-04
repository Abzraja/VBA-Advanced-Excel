Sub VBAChallenge():


Dim ws As Worksheet
Dim Ticker As String
Dim LastRow As Long
Dim YearOpen As Double
Dim YearClose As Double
Dim YearChange As Double
Dim YearPercentChange As String
Dim TotalStockVolume As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As String


For Each ws In Worksheets


' Keep track of the location for each ticker letter in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set value of first open
YearOpen = ws.Cells(2, 3).Value
        
'Print value of year open
'Range("O" & Summary_Table_Row).Value = YearOpen
        
'Set TotalStockVolume
TotalStockVolume = 0
    
    
    'For all the filled rows
    For i = 2 To LastRow
        
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        'If the Next row value is different to the previous
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


            'Set the ticker letter
            Ticker = ws.Cells(i, 1).Value
        
       
            'Year End Value
            YearClose = ws.Cells(i, 6).Value
        
            
            'Print Ticker Header
            ws.Range("I1").Value = "Ticker"
            
            'Print Yearly Change Header
            ws.Range("J1").Value = "Yearly Change"
            
            'Print PercentChange Header
            ws.Range("K1").Value = "Percent Change"
            
            'Print Total Stock Volume Header
            ws.Range("L1").Value = "Total Stock Volume"
            
            'Print the Ticker Letter in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
        
            'Print value of year open to chek if correct - this is just for testing
            'Range("O" & Summary_Table_Row).Value = YearOpen
        
       
            'Print the YearClose Value to check if correct - this is just for testing
            'Range("P" & Summary_Table_Row).Value = YearClose
        
            'Calculate the yearly change
            YearChange = YearClose - YearOpen
            
        
            'Print Yearly Change
            ws.Range("J" & Summary_Table_Row).Value = YearChange
            
                'Test IF cell is greater or less than 0
                If YearChange > 0 Then
                    
                    'If Value > 0 Color Cell Green
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                'Else color cell red
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
        
                'Avoid divisible by 0 overflow error
                If YearChange <> 0 And YearOpen <> 0 Then
            
                    'Calculate the percent change between open and close
                    YearPercentChange = YearChange / YearOpen
                    
                End If
        
            'Print Percentage Change
            ws.Range("K" & Summary_Table_Row).Value = YearPercentChange
            'Format as Percentage
            ws.Range("K" & Summary_Table_Row).Value = FormatPercent(ws.Range("K" & Summary_Table_Row))
        
            'Print Total Stock Volume
            ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
         
            'Reset TotalStockVolume
            TotalStockVolume = 0
        
        
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
            'Set value of next open
            YearOpen = ws.Cells(i + 1, 3).Value
        
       

      
        End If


    Next i
    
    
'Print Bonus Ticker Header
ws.Range("P1").Value = "Ticker"
            
'Print Bonus Value Header
ws.Range("Q1").Value = "Value"

'Print Bonus Greatest Columns
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
    
'Find Min and Max Values
ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row - 1))
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row - 1))
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row - 1))

'Read Greatest Values into Variables
GreatestIncrease = ws.Range("Q2").Value
GreatestDecrease = ws.Range("Q3").Value
GreatestVolume = ws.Range("Q4").Value


    'Iterate through Summary Table
    For j = 2 To (Summary_Table_Row - 1)
    
        'If condition to find row for Greatest Increase
        If ws.Cells(j, 11).Value = GreatestIncrease Then
        
        'Print ticker value for Greatest Increase
        ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
    
        End If
        
        
        'If condition to find row for Greatest Decrease
        If ws.Cells(j, 11).Value = GreatestDecrease Then
        
        'Print ticker value for Greatest Decrease
        ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
    
        End If
        
        
        'If condition to find row for Greatest Volume
        If ws.Cells(j, 12).Value = GreatestVolume Then
        
        'Print ticker value for Greatest Volume
        ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
    
        End If
        
        
        
    'End summary table iteration
    Next j
    
    
    
'Format as percentage
ws.Range("Q2").Value = FormatPercent(ws.Range("Q2"))
ws.Range("Q3").Value = FormatPercent(ws.Range("Q3"))
    
    
    
'Autofit columns so data displayed correctly
ws.Columns("A:Q").AutoFit
    
    
Next ws
        

End Sub






