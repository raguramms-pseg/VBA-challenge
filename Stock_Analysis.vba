Attribute VB_Name = "Module1"

'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.


Sub Stock_Analysis_Multiple_Year()
    Dim Ticker As String
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Stock_Volume As LongLong
    Dim Year_Opening_Price As Double
    Dim Year_Closing_Price As Double
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As LongLong
    
        
    Dim Sheet_Count As Integer
    Dim I As Integer
    Dim J As Long
    Dim RowIndex As Integer
    Dim lastrow As Long
    
    Sheet_Count = ActiveWorkbook.Worksheets.Count
    ' Begin the loop.
     For I = 1 To Sheet_Count
        ActiveWorkbook.Worksheets(I).Range("I1").Value = "Ticker"
        ActiveWorkbook.Worksheets(I).Range("J1").Value = "Yearly Change"
        ActiveWorkbook.Worksheets(I).Range("K1").Value = "Percent Change"
        ActiveWorkbook.Worksheets(I).Range("L1").Value = "Total Stock Volume"
        
        lastrow = ActiveWorkbook.Worksheets(I).Cells(Rows.Count, 1).End(xlUp).Row
                
        ' Initially set the totals to be 0 for each Ticker
        
        Yearly_Change = 0
        Percentage_Change = 0
        Total_Stock_Volume = 0
        Ticker = ActiveWorkbook.Worksheets(I).Range("A2").Value
        Year_Opening_Price = ActiveWorkbook.Worksheets(I).Range("C2").Value
        Year_Closing_Price = 0
        Greatest_Percent_Increase = 0
        Greatest_Percent_Decrease = 0
        Greatest_Total_Volume = 0
        
        RowIndex = 2
        
        ' Loop through each row in the worksheet and accumulate totals for each ticker
        For J = 2 To lastrow
            If ActiveWorkbook.Worksheets(I).Cells(J, 1).Value = Ticker Then
                Total_Stock_Volume = Total_Stock_Volume + ActiveWorkbook.Worksheets(I).Cells(J, 7).Value
                
            Else
                Year_Closing_Price = ActiveWorkbook.Worksheets(I).Cells(J - 1, 6).Value
                Yearly_Change = Year_Closing_Price - Year_Opening_Price
                Percentage_Change = (Yearly_Change / Year_Opening_Price)
                If Percentage_Change > Greatest_Percent_Increase Then
                    Greatest_Percent_Increase = Percentage_Change
                    Greatest_Increase_Ticker = Ticker
                End If
                If Percentage_Change < Greatest_Percent_Decrease Then
                    Greatest_Percent_Decrease = Percentage_Change
                    Greatest_Decrease_Ticker = Ticker
                End If
                               
                ActiveWorkbook.Worksheets(I).Cells(RowIndex, 9) = Ticker
                ActiveWorkbook.Worksheets(I).Cells(RowIndex, 10) = Yearly_Change
                If Yearly_Change > 0 Then
                ' Set the Cell Colors to Green
                    ActiveWorkbook.Worksheets(I).Cells(RowIndex, 10).Interior.ColorIndex = 4
                Else
                ' Set the Cell Colors to Red
                    ActiveWorkbook.Worksheets(I).Cells(RowIndex, 10).Interior.ColorIndex = 3
                End If
                
                ActiveWorkbook.Worksheets(I).Cells(RowIndex, 11) = FormatPercent(Percentage_Change)
                ActiveWorkbook.Worksheets(I).Cells(RowIndex, 12) = Total_Stock_Volume
                If Total_Stock_Volume > Greatest_Total_Volume Then
                    Greatest_Total_Volume = Total_Stock_Volume
                    Greatest_Total_Ticker = Ticker
                End If
                
                RowIndex = RowIndex + 1
                
                Ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
                Yearly_Change = 0
                Percentage_Change = 0
                Total_Stock_Volume = ActiveWorkbook.Worksheets(I).Cells(J, 7).Value
                Year_Opening_Price = ActiveWorkbook.Worksheets(I).Cells(J, 3).Value
                
            End If
    
        Next J
        
        Year_Closing_Price = ActiveWorkbook.Worksheets(I).Cells(lastrow, 6).Value
        Yearly_Change = Year_Closing_Price - Year_Opening_Price
        Percentage_Change = (Yearly_Change / Year_Opening_Price)
        
        If Percentage_Change > Greatest_Percent_Increase Then
            Greatest_Percent_Increase = Percentage_Change
            Greatest_Increase_Ticker = Ticker
        End If
        If Percentage_Change < Greatest_Percent_Decrease Then
            Greatest_Percent_Decrease = Percentage_Change
            Greatest_Decrease_Ticker = Ticker
        End If
                
        ActiveWorkbook.Worksheets(I).Cells(RowIndex, 9) = Ticker
        ActiveWorkbook.Worksheets(I).Cells(RowIndex, 10) = Yearly_Change
        
        If Yearly_Change > 0 Then
        ' Set the Cell Colors to Green
            ActiveWorkbook.Worksheets(I).Cells(RowIndex, 10).Interior.ColorIndex = 4
        Else
        ' Set the Cell Colors to Red
            ActiveWorkbook.Worksheets(I).Cells(RowIndex, 10).Interior.ColorIndex = 3
        End If
        
        ActiveWorkbook.Worksheets(I).Cells(RowIndex, 11) = FormatPercent(Percentage_Change)
        ActiveWorkbook.Worksheets(I).Cells(RowIndex, 12) = Total_Stock_Volume
        
        If Total_Stock_Volume > Greatest_Total_Volume Then
            Greatest_Total_Volume = Total_Stock_Volume
            Greatest_Total_Ticker = Ticker
        End If
        
        ActiveWorkbook.Worksheets(I).Range("P1") = "Ticker"
        ActiveWorkbook.Worksheets(I).Range("Q1") = "Value"
        
        ActiveWorkbook.Worksheets(I).Range("O2") = "Greatest % Increase"
        ActiveWorkbook.Worksheets(I).Range("P2") = Greatest_Increase_Ticker
        ActiveWorkbook.Worksheets(I).Range("Q2") = FormatPercent(Greatest_Percent_Increase)
        ActiveWorkbook.Worksheets(I).Range("Q2").Interior.ColorIndex = 4
        
        ActiveWorkbook.Worksheets(I).Range("O3") = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(I).Range("P3") = Greatest_Decrease_Ticker
        ActiveWorkbook.Worksheets(I).Range("Q3") = FormatPercent(Greatest_Percent_Decrease)
        ActiveWorkbook.Worksheets(I).Range("Q3").Interior.ColorIndex = 3
              
        ActiveWorkbook.Worksheets(I).Range("O4") = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(I).Range("P4") = Greatest_Total_Ticker
        ActiveWorkbook.Worksheets(I).Range("Q4") = Greatest_Total_Volume
        

    Next I
    
End Sub



