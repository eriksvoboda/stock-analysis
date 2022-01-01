# Stock Analysis

## Overview of Project

  Steve wants to understand how different stocks perform in 2017 and 2018. The purpose of this project is to provide Steve the ability to analyze all stocks in 2017 and 2018. Additionally this project was focused on refactoring the code previously used to analyze the stocks in a given year.

## Results

### Stock Return Results
  In 2017 the top three stocks with the largest Return, in order from greatest to least were; DQ (199.4%), SEDG (184.5%), and ENPH (129.5%). The three stocks with the worst Return in 2017, in order from least to greatest were; TERP (-7.2%), RUN (5.5%), and AY (8.9%). The top three stocks in 2018, from greatest to least were; RUN (84%), ENPH (81.9%), and VSLR (-3.5%). The worst performing 2018 stocks in terms of Return, in order from least to greatest were; DQ(-62.6%), JKS (-60.5), and SPWR (-44.6%). The average Return of all stocks in 2017 was 67.3% and in 2018 the all stock average was -8.5%.
  
  Based on the Returns from 2017 and 2018 ENPH is the stock Steve should suggest to his parents as it has the highest average Return of 105.7% over the two years and outperforms the average of all stocks in both years. DQ has an average return of 68.4% and while it beats the average Return of all stocks in 2017 it's worse to the all stock average in 2018. DQ is a more volatile stock than ENPH so Steve should suggest ENPH to his parents. 
  
### Results of Refactoring Code  

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
   tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'The instructions say only do it for tickerVolumes but it wasn't working so I tried adding tickerStartingPrices and tickerEndingPrices and it at least didn't break the code so I kept them all
            For i = 0 To 11
                tickerVolumes(i) = 0
                tickerStartingPrices(i) = 0
                tickerEndingPrices(i) = 0
             Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        'Found other code on Stack Overflow https: //stackoverflow.com/questions/69684996/runtime-6-overflow-error-refactoring-code-for-stock-analysis because I was getting an Overflow error but now this original piece of code works.
  tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        'End If
         End If
         
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
       
            '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        'End If
        End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'This finally works! I don't know what i did differently but it works, I just retyped the same code that I already had
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
    
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

### Run time difference of Original Code and Refactored Code by year

Run time for 2017 Analysis: Original

![](/Resources/Original_2017_Runtime.png)


Run time for 2017 Analysis: Refactored

![](/Resources/VBA_Challenge_2017.png)

Run time for 2018 Analysis: Original

![](/Resources/Original_2018_Runtime.png)

Run time for 2018 Analysis

![](/Resources/VBA_Challenge_2018.png)

## Summary & Conclusions

  One way this code could be improved further is including 2017 and 2018 in one pull. Then Steve would be able to compare Total Daily Volume and Return of 2017 and 2018 side by side versus having to pull each year individually.
