Attribute VB_Name = "Module1"
Sub StockData()


Dim ws As Worksheet

' loop through each worksheet.
For Each ws In Worksheets

    'Declaring variables
    Dim CurrentTicker As String
    Dim FollowingTicker As String
    Dim Day As Date
    Dim FirstDay As Date
    Dim LastDay As Date
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim HighPrice As Double
    Dim LowPrice As Double
    Dim Volume As Double
    Dim lastrow As Double
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim counter As Integer
    Dim ColorGreen As Integer
    Dim ColorRed As Integer
    Dim Greatest_percent_increase As Variant
    Dim Greatest_percent_decrease As Variant
    Dim Percent_increase_ticker As String
    Dim Percent_decrease_ticker As String
    Dim Greatest_volume As Double
    Dim Greatest_volume_ticker As String
    
    'Titling new table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    '-----------------------------------------------------------------------
    'BONUS
    'Titling rows in second table
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % decrease"
    ws.Cells(4, 14).Value = "Greatest total volume"
    
    'Titling columns in second table
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    '-----------------------------------------------------------------------
    
    'Setting initial values
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'calculating last row
    counter = 2
    counter2 = 2
    OpenPrice = ws.Cells(2, 3).Value
    Volume = 0
    ColorGreen = 4
    ColorRed = 3
    
    'Calculcating change in rows
    For Current = 2 To lastrow
        following = Current + 1
        CurrentTicker = ws.Cells(Current, 1).Value
        FollowingTicker = ws.Cells(following, 1).Value
        Volume = Volume + ws.Cells(Current, 7).Value
    
       
        If CurrentTicker <> FollowingTicker Then
            'Assign the ticker value to the new table
            ws.Cells(counter, 9).Value = CurrentTicker
    
            'Calculating yearly change
            ClosePrice = ws.Cells(Current, 6).Value
            YearlyChange = ClosePrice - OpenPrice
            ws.Cells(counter, 10).Value = YearlyChange
            
            'Conditional formatting
            If ws.Cells(counter, 10).Value >= 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = ColorGreen
            ElseIf ws.Cells(counter, 10).Value < 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = ColorRed
                  
            End If
            
            'Calculation percent change
            If OpenPrice <> 0 Then
                PercentChange = YearlyChange / OpenPrice
                ws.Cells(counter, 11).Value = Format(PercentChange, "Percent")
            ElseIf OpenPrice = 0 Then
                ws.Cells(counter, 11).Value = "Invalid"
            
            End If
            
            'Assign Volume to the table
            ws.Cells(counter, 12).Value = Volume
            
    
            'Reset values for next ticker
            Volume = 0
            OpenPrice = ws.Cells(following, 3).Value
            counter = counter + 1
            
            
            
        End If
    
    Next Current
    
    '------------------------------------------------------------------------
    'BONUS
    
    Greatest_percent_increase = 0
    Greatest_percent_decrease = 0
    Greatest_volume = 0
    
    For Current = 2 To lastrow:
        following = Current + 1
    
    
        If IsNumeric(ws.Cells(Current, 11).Value) Then
            
        'Calculate greatest % increase
            If ws.Cells(Current, 11).Value > Greatest_percent_increase Then
            Greatest_percent_increase = ws.Cells(Current, 11).Value
            Percent_increase_ticker = ws.Cells(Current, 9).Value
            End If
            
        'Calculate greatest % decrease
            If ws.Cells(Current, 11).Value < Greatest_percent_decrease Then
                Greatest_percent_decrease = ws.Cells(Current, 11).Value
                Percent_decrease_ticker = ws.Cells(Current, 9).Value
            End If
            
        'Calculate greatest total volume
            If ws.Cells(Current, 12).Value > Greatest_volume Then
                Greatest_volume = ws.Cells(Current, 12).Value
                Greatest_volume_ticker = ws.Cells(Current, 9).Value
            End If
        End If
       
    Next Current
        
    'Assigning values to new table
    ws.Cells(2, 16).Value = Format(Greatest_percent_increase, "Percent")
    ws.Cells(3, 16).Value = Format(Greatest_percent_decrease, "Percent")
    ws.Cells(2, 15).Value = Percent_increase_ticker
    ws.Cells(3, 15).Value = Percent_decrease_ticker
    ws.Cells(4, 15).Value = Greatest_volume_ticker
    ws.Cells(4, 16).Value = Greatest_volume

Next

End Sub


