Attribute VB_Name = "ModuleMYSD"
Sub Multiple_year_stock_data()

 ' Create variables
    Dim i As Long
    Dim ws As Worksheet

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

    ' Create variables
    Dim Last_Row As Long
    Dim Ticker As String
    Dim Quarterly_Change As Double
  
    Dim Percent_Change As Double
  
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    Dim closeprice As Double
    Dim openprice As Double
    
    Dim cpp As Double
    Dim opp As Double
    
    ' Determine the Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Keep track of the location for each Ticker Symbol category in the Summary Table
    Dim Summary_Table_Rows As Integer
    Summary_Table_Rows = 2
    
     ' Label Columns of Summary Table
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Quarterly Change"
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'Set first openprice for Quarterly_Change & Percent_Change
    openprice = ws.Cells(2, 3).Value
    opp = ws.Cells(2, 3).Value
    
    'Loop through all the Ticker symbols
    For i = 2 To Last_Row
    
    'Check if we are still within the same Ticker symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set the Ticker & closeprice for Quarterly_Change & Percent_Change
        Ticker = ws.Cells(i, 1).Value
        closeprice = ws.Cells(i, 6).Value
        cpp = ws.Cells(i, 6).Value
         
        'Add to the 'Quarterly Change' Total
        Quarterly_Change = closeprice - openprice
        openprice = ws.Cells(i + 1, 3).Value
    
        'Add to the 'Percent Change' Total
        Percent_Change = ((cpp - opp) / opp)
        opp = ws.Cells(i + 1, 3).Value
    
        'Add to the 'Total Stock Volume' Total
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        'Print 'Ticker Symbol' in the summary table
        ws.Range("I" & Summary_Table_Rows).Value = Ticker
        
        'Print 'Quarterly Change' in the Summary Table
        ws.Range("J" & Summary_Table_Rows).Value = Quarterly_Change
        ws.Range("J" & Summary_Table_Rows).NumberFormat = "0.00"
        
        'Print the 'Percent Change' in the Summary Table
        ws.Range("K" & Summary_Table_Rows).Value = Percent_Change
        ws.Range("K" & Summary_Table_Rows).NumberFormat = "0.00%"

        'Print the 'Total Stock Volume' in the Summary Table
        ws.Range("L" & Summary_Table_Rows).Value = Total_Stock_Volume
        
        'Add one to the summary table row
        Summary_Table_Rows = Summary_Table_Rows + 1
        
        'Reset the 'Total Stock Volume Total
        Total_Stock_Volume = 0
             
        Else
        
            'Add to the 'Total Stock Volume' Total
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
         

    End If

        Next i
    
    'Change interior color of 'Quarterly Change" Column to "Green for Positive / Red for Negative
    For i = 2 To Last_Row
        For j = 10 To 10
        
    ' Create variables to hold Last_ Row, and determine the Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Set New Quartley_Change value
    Set NewQC = ws.Cells(i, j)
        
    If NewQC.Value < 0 Then
        
        'Color Negative cells Red
        ws.Cells(i, j).Interior.ColorIndex = 3
        
        ElseIf NewQC.Value > 0 Then
        
        'Color Positive cells Green
        ws.Cells(i, j).Interior.ColorIndex = 4

        End If
        
        Next j
    Next i
    
    
     ' Label Columns & Rows for "Greatest" Summary Table
     ws.Cells(1, 15).Value = "Ticker"
     ws.Cells(1, 16).Value = "Value"
     ws.Cells(2, 14).Value = "Greatest % Increase"
     ws.Cells(3, 14).Value = "Greatest % Decrease"
     ws.Cells(4, 14).Value = "Greatest Total Volume"

    'Create variables for Greatest % Increase/Decrease
    
    Dim firstincreasevalue As Double
    Dim greatincreaseticker As String
    Dim greatincrease As Double
   
    
    Dim firstdecreasevalue As Double
    Dim greatdecreaseticker As String
    Dim greatdecrease As Double
    
    'Initialize the increase/decrease value
    greatincrease = -9999999
    greatdecrease = 9999999

    
    ' Create variables to hold Last_ Row, and determine the Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through all the "Percent Change" values
    For i = 2 To Last_Row
    For j = 11 To 11

    'Set the first increase/decrease "Percent Change" value
    firstincreasevalue = ws.Cells(i + 1, j).Value
    firstdecreasevalue = ws.Cells(i + 1, j).Value
  
        'Check for greatest increase/decrease value, if it is not...
        If firstincreasevalue > greatincrease Then
    
        greatincrease = firstincreasevalue
        greatincreaseticker = ws.Cells(i + 1, j - 2).Value
    
        'Print Greatest % Increase value in "Greatest" Summary Table
        ws.Range("p2") = greatincrease
        ws.Range("p2").NumberFormat = "0.00%"
        
        'Print Greatest % Increase Ticker in "Greatest" Summary Table
        ws.Range("o2") = greatincreaseticker
    
    
    ElseIf firstdecreasevalue < greatdecrease Then
           greatdecrease = firstdecreasevalue
           greatdecreaseticker = ws.Cells(i + 1, j - 2).Value
       
           'Print Greatest % Decrease value in "Greatest" Summary Table
           ws.Range("p3") = greatdecrease
           ws.Range("p3").NumberFormat = "0.00%"
           
           'Print Greatest % Decrease Ticker in "Greatest" Summary Table
           ws.Range("o3") = greatdecreaseticker
    
     
        End If
        
        Next j
    Next i

    'Create variable for Greatest Total Volume
    Dim firstgreatvolume As Double
    Dim greatvolumeticker As String
    Dim greatvolume As Double

    'Initialize Greatest Total Volume value
    greatvolume = -999999999
    
    ' Create variables to hold Last_ Row, and determine the Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through all the "Total Stock Volume" values
    For i = 2 To Last_Row
        For j = 12 To 12
    
    'Set the first Total Stock Volume value
    firstgreatvolume = ws.Cells(i + 1, j).Value
  
    
        'Check for greatest total volume value, if it is not...
        If firstgreatvolume > greatvolume Then
    
        greatvolume = firstgreatvolume
        greatvolumeticker = ws.Cells(i + 1, j - 3).Value
    
        'Print Greatest Total Volume value in "Greatest" Summary Table
        ws.Range("p4") = greatvolume
        
        'Print Greatest Total Volume Ticker in "Greatest" Summary Table
        ws.Range("o4") = greatvolumeticker
    
        End If
        
        Next j
    Next i

    Next ws
End Sub

