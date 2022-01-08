## Stock Analysis with VBA
# Overview of Project
In module 2 Steve wanted more information on specific stocks from 2017 and 2018 to determine if investment was approrpiraete for his parents. The code created in Module 2 provided Steve with the desired information for the selected stocks in those two years. Steve wanted to expand the code to evaluate all stocks, therefore this challenge had us edit or refactor the code from Module 2. The refactoring is beneficial as it can support large data sets, increases the efficiency of the code, and is easier for other coders to read, understand and replicate. 
# Results
Before refactoring I copied the code provided by the challenge. I copied this code into the new Sub and used the numbers to guide me through the refactoring. The concluding code is as follows:

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
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
             tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) = Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                    
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
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

The results of this code showed that the stocks performed better in 2017 than 2018 as shown below. 

![StockResults_2017](https://user-images.githubusercontent.com/94129215/148661566-bff7b671-bdc9-46c7-b75b-3ac3504992b5.png)
![StockResults_2018](https://user-images.githubusercontent.com/94129215/148661567-8f329d73-81f3-4b50-baf8-7c1e861029fd.png)

The refactoring was successful as the performance time decreased significantly. As you can see from the images below, the speed went from almost three tenths of a second to under one tenth of a second. 
*Module 2 Code Run Time
<img width="266" alt="greenstock2017" src="https://user-images.githubusercontent.com/94129215/148661713-9ab71703-cd7a-457d-8b3b-57745f676e1a.png">
<img width="262" alt="greenstock2018" src="https://user-images.githubusercontent.com/94129215/148661718-01da5242-2f05-4dd3-910a-644a126bfe6e.png">
*Refactored Code Run-Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/94129215/148661415-41cf57aa-e7f2-4c97-a9ea-088d69914cb4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/94129215/148661647-151479ef-49a8-4d19-9e7f-4c9a4a770344.png)


# Summary
*Advantages/Pros and Disadvantages/Cons to Refactoring* 
The main advantage to refactoring is the speed in which the macro can be run. It cut the time dramatically. The quicker the macro can run, the quicker clients can get the information for which they are looking. In addition, refactoring improves the design of software by making it easier to understand, debug and program faster. Some of the disadvanges are the risk becomes too great when the application is too big or do do not have proper test cases you can not trial your refactoring. It can make the application disfuctional. 

