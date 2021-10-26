# stock-analysis with VBA


## Overview of Project

Steve is planning to do more research for his parents. He wants to expand the dataset to include the entire stock market.
The stock analysis VBA code we developed in module 2 might not work as well and take a long time to excecute for a larger amount of stocks.

Therefore the purpose of the project is to restructure the existing code in a way that improves the internal structure but doesn't change its external behavior.
We aren’t adding new functionality, but want to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read

## Results

### Analysis

#### First step: Using the starter code in Visual Basic Editor.
I reused the existing code to set up the Timer, create the user InputBox, activate the "All Stocks Analysis" worksheet, add the header, header row and initialize the array of tickers with the ticker values, activate the yearValue worksheet, and get the number of rows in the sheet to loop over.

  ![image](https://user-images.githubusercontent.com/91682586/138772963-ebb74438-ba95-406b-a745-85f50c30bb10.png)


#### 1a) I then created the tickerIndex that will later be used to access the arrays created in Step 1b) and set it to 0. 
    '1a) Create a ticker Index
    
    tickerIndex = 0

#### 1b) I created three output arrays for ticker Volumes, tickerStarterPrices, and tickerEndingPrices. The Datatype for ticker was set to Long, tickerStartingPrice and tickerEndingPrice were set to Single. 
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
#### 2a) I created a for loop to initialize the tickerVolumes to zero.
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To RowCount 
    
        tickerVolumes(tickerIndex) = 0
               
    Next i
   
    

#### ![#f03c15](https://via.placeholder.com/15/f03c15/000000?text=+) 2b) Created code that will loop over all the rows in the spreadsheet. 
            '2b) Loop over all the rows in the spreadsheet.
            
            For i = 2 To RowCount
    
  ##### 3a) This code increases the volume for the current ticker. 
            '3a) Increase volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value   
            
  ##### 3b) This code checks if the current row is the first row with the selected tickerIndex and if true it assigns the current starting price to the tickerStartingPrice  variable.
            '3b) Check if the current row is the first row with the selected tickerIndex. If true assign starting price.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            'Assign Starting Price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
                     
   ##### 3c) This code checks if the current row is the last row with the selected ticker. If true it assigns the current ending Price to the tickerEndingPrice variable. If the next row’s ticker doesn’t match, the tickerIndex gets increased. 
   
            '3c) check if the current row is the last row with the selected ticker                        
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            'Assign Ending Price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
   ##### 3d) This code increases the tickerIndex.
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            tickerIndex = tickerIndex + 1
                                    
            End If
            
  #### ![#f03c15](https://via.placeholder.com/15/f03c15/000000?text=+) The loop moves on to the next row.
        Next i    
   
       
   #### 4) This code loops through our arrays to output the Ticker, Total Daily Volume, and Return.
        'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
        Next i 
  
  #### Formatting - I reused the existing Code to activate and format the Output worksheet, end the timer, and finish the Macro with End Sub. 
  
   'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
         columns("B").AutoFit

         dataRowStart = 4
         dataRowEnd = 15

        For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    'end timer    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub
  
  
### Summary

- Through refactoring code becomes easier to understand or read, easier to to update and improve. This can save time and money in the future. 
- It helps the author coming back to read the code after a while as well as outside users.
- Refactoring can make the code more flexibel for other uses. 
- A disadvantage is that with complex code it might not be clear from the beginning how long exactly the completioin of the process might take and if there is a solution at all.   Because of the complexity you might end up spending a lot of time with little improvement in the end.
- For our refactored stock-analysis code the main advantage is that it runs much faster. For 2018 the original code needed 0.9335938 seconds to run versus a run time of        0.1367188 for the refactored code. For 2017 the original code needed 0.8359375 seconds versus the refactored code taking 0.109375 seconds.
- A faster calculation is important for larger datasets.  
- The refactored code can be reused for other projects that require looping over items. 
- The new code is a bit more complex than the original code, e.g. requires understanding of arrays. 
