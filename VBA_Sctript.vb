Sub StocksAnalysys()

'Declare all needed variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
Dim OpenPrice_BY As Double
Dim ClosePrice_EY As Double
Dim Summary_Table_Row As Long
Dim Last_Row As Long
Dim Max_Change As Double
Dim Min_Change As Double
Dim Max_Tot_Volume As Double

' Loop through all sheets
Dim WS As Worksheet
For Each WS In Worksheets

    'Determine last row in original Ticker Table
    Last_Row = WS.Cells(Rows.Count, 1).End(xlUp).Row

   'Sort data Ascending on ticker and Date
    WS.Sort.SortFields.Clear
    WS.Sort.SortFields.Add Key:=Range("A2:A" & Last_Row), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    WS.Sort.SortFields.Add Key:=Range("B2:B" & Last_Row), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With WS.Sort
        .SetRange Range("A1:G" & Last_Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    'Name new columns
    WS.Range("I1") = "Ticker"
    WS.Range("J1") = "Yearly_Change"
    WS.Range("K1") = "Percent_Change"
    WS.Range("L1") = "Total_Stock_Volume"
        
    'start populatig summary table
    Summary_Table_Row = 2
    Total_Stock_Volume = 0
    OpenPrice_BY = WS.Cells(2, 3).Value
        
    For I = 2 To Last_Row
        
        If WS.Cells(I + 1, 1).Value <> WS.Cells(I, 1).Value Then
        
        'Populate list of unique tickers
        Ticker = WS.Cells(I, 1).Value
        WS.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Calculate Yearly Change and Percent Change and record results in summary table
        ClosePrice_EY = WS.Cells(I, 6).Value
        Yearly_Change = ClosePrice_EY - OpenPrice_BY
        WS.Range("J" & Summary_Table_Row).Value = Yearly_Change
            If OpenPrice_BY = 0 Then
                Percent_Change = 0
                Else
                Percent_Change = Yearly_Change / OpenPrice_BY
            End If
            WS.Range("K" & Summary_Table_Row).Value = Percent_Change
             
             'Format color of Percentage change cell
            If Percent_Change > 0 Then
                WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Percent_Change < 0 Then
                WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 1
                WS.Range("K" & Summary_Table_Row).Font.ColorIndex = 2
            End If
        
        'Calculate Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + WS.Cells(I, 7).Value
        WS.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'Next row in summary table
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Stock_Volume = 0
        OpenPrice_BY = WS.Cells(I + 1, 3).Value
        
        Else
        Total_Stock_Volume = Total_Stock_Volume + WS.Cells(I, 7).Value
    
        End If
        
    Next I

    ' Formating Column width, format number
    WS.Columns("I:L").AutoFit
    WS.Columns("K").NumberFormat = "0.00%"
    WS.Columns("L").NumberFormat = "#,###"

    'create 2nd summary table
    WS.Range("O2").Value = "Greatest % Increase"
    WS.Range("O3").Value = "Greatest % Decrease"
    WS.Range("O4").Value = "Greatest Total Volume"
    WS.Range("P1").Value = "Ticker"
    WS.Range("Q1").Value = "Value"
    
       'Determine last row in the 1st summary table
    Last_Row = WS.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' finding Greatest % Increase, Greatest % Decrease and greatest Total Volume
    Max_Change = WorksheetFunction.Max(WS.Range("K2:K" & Last_Row))
    Min_Change = WorksheetFunction.Min(WS.Range("K2:K" & Last_Row))
    Max_Tot_Volume = WorksheetFunction.Max(WS.Range("L2:L" & Last_Row))
        
        For I = 2 To Last_Row
            If WS.Cells(I, 11) = Max_Change Then
                Ticker = WS.Cells(I, 9).Value
                WS.Range("Q2").Value = Max_Change
                WS.Range("P2").Value = Ticker
                
            ElseIf WS.Cells(I, 11) = Min_Change Then
                Ticker = WS.Cells(I, 9).Value
                WS.Range("Q3").Value = Min_Change
                WS.Range("P3").Value = Ticker
                      
            End If
            
            If WS.Cells(I, 12) = Max_Tot_Volume Then
                Ticker = WS.Cells(I, 9).Value
                WS.Range("Q4").Value = Max_Tot_Volume
                WS.Range("P4").Value = Ticker
            End If
            
        Next I
        WS.Range("Q2:Q3").NumberFormat = "0.00%"
        WS.Range("Q4").NumberFormat = "#,###"
        WS.Columns("O:Q").AutoFit
        
Next WS

End Sub


