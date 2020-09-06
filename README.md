# VBA-challenge

VBA has been considered one of most important tools to automate repetitive word- and data-processing functions, and generate custom forms, graphs, and reports
Under this project, I code a VBA script used to analyze real stock market data in a Microsoft Excel workbook as well as re-factoring script to analyze a thousands of stocks rather than just an dozen of stocks 

About script

You can find the script inside the VBA Stocks folder of this repository. 
The script file is called AllStocksAnalysis. After you download and open up the multiple year stock data Excel workbook, you can run the script by doing the following: 

    •	Click the Developer tab. 
    •	Click Visual Basic to open the Visual Basic editor. Inside the Visual Basic editor, click File > Import File and import the AllStocksAnalysis
    •	Open up the AllStocksAnalysis file in the Visual Basic editor and then click the Run Macro button (green play icon) in the toolbar to run the script. Tthe script does take some time to run because it is running on every sheet. So, no need to run it more than once. As the script runs, it is doing the             following: 
    •	It loops through all the stocks for one year for each run and takes the following information: The ticker symbol 
           o Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
           o The percent change from opening price at the beginning of a given year to the closing price at the end of that year. 
           o The total stock volume of the stock. 
    •	It applies conditional formatting by highlighting positive yearly change values in green and negative yearly change values in red. 

Sample Output
After the script has completed, go to the Excel workbook, and you should see the results of the script. Here are screenshots of what the output looks like when I ran the scripts on my computer. These screenshots are also available in the VBAStocks/screenshots folder of this repository.

<img width="222" alt="Screen Shot 2020-09-05 at 8 25 05 PM" src="https://user-images.githubusercontent.com/70412047/92330805-e8f1fb80-f03f-11ea-9b04-9b005d60dc7c.png">
<img width="223" alt="Screen Shot 2020-09-05 at 8 24 48 PM" src="https://user-images.githubusercontent.com/70412047/92330809-ee4f4600-f03f-11ea-87ec-a8f1482854e4.png">


  
Refactoring script

As we might be notified that when analyzing 12 stocks, loop through all stocks three times wont be big problems but It absolutely will be big difference if we are trying to analyze the entire stock market.

In order to improve efficiency, we will loop through all stocks one time and append data into an array so that we can populate data we need to. 

    •	We create an empty array to instore all unique tickerIndex and it will be used to populate total volume, start price and end price later. 
    •	Set up three more arrays to contain total volume, start price, and end price so that we will be able to populate data. 

Refactor code:

Here is the code and I will also upload that file to this GitHub repository.

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

   
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over and intialized of dynamic arrays for tickers
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    Dim tickers() As String
    Index = 0
    ReDim tickers(Index)
    For i = 2 To RowCount
         If IsError(Application.Match(Cells(i, 1).Value, tickers, False)) Then
            ReDim Preserve tickers(Index)
            tickers(Index) = Cells(i, 1).Value
            Index = Index + 1
         End If
    Next i
    
    '1a) Create a ticker Index
         tickerIndex = 0
    

    '1b) Create three output arrays
         Dim totalVolumes(11) As Long
         Dim StartPrice(11) As Single
         Dim EndPrice(11) As Single

         
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
          tickerIndex = 0
          priceIndex = 0
          totalVolumes(tickerIndex) = 0
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
             totalVolumes(tickerIndex) = totalVolumes(tickerIndex) + Cells(i, 8)
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i, 1) = Cells(i + 1, 1) And Cells(i - 1, 1) <> Cells(i, 1) Then
              StartPrice(priceIndex) = Cells(i, 3).Value
              priceIndex = priceIndex + 1
            
         End If
        
        '3c) check if the current row is the last row with the selected ticker
           If Cells(i, 1) <> Cells(i + 1, 1) Or Cells(i + 1, 1) = "" Then
            '3d Increase the tickerIndex.
                EndPrice(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
        
           End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
           Cells(4 + i, 1).Value = tickers(i)
           Cells(4 + i, 2).Value = totalVolumes(i)
           If StartPrice(i) = 0 Then
           Cells(4 + i, 3).Value = 0
           Else
           Cells(4 + i, 3).Value = (EndPrice(i) - StartPrice(i)) / StartPrice(i)
          End If
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



Time Output:

Here is output that we can see the time difference if we apply array instead of loop through array three times. 

Original time: 

<img width="427" alt="Screen Shot 2020-09-05 at 8 17 12 PM" src="https://user-images.githubusercontent.com/70412047/92330921-b268b080-f040-11ea-9d17-c9f07e21c792.png">
<img width="425" alt="Screen Shot 2020-09-05 at 8 16 52 PM" src="https://user-images.githubusercontent.com/70412047/92330855-2d7d9700-f040-11ea-9297-ac8ff1e976ba.png">
 
 
Refactor time:

<img width="421" alt="Screen Shot 2020-09-05 at 8 24 26 PM" src="https://user-images.githubusercontent.com/70412047/92330850-2c4c6a00-f040-11ea-8547-4711e682239f.png">
<img width="422" alt="Screen Shot 2020-09-05 at 8 19 12 PM" src="https://user-images.githubusercontent.com/70412047/92330853-2ce50080-f040-11ea-9a6e-21e0d7bbb1e6.png">
VBA Challenge
 
 

Advantage and disadvantage of in the refactoring code
       •  Running time is shorter and improve efficiency 
       •  May need more time to re-editing and test script




Advantage and disadvantage of In the original code
       •  Easy to follow and code it down
       •. It is not efficient in terms of running time



