# stock-analysis with VBA
click here for the excel file: 
## Overview of Project

Steve is planning to do more research for his parents. He wants to expand the dataset to include the entire stock market.
The stock analysis VBA code we developed in module 2 might not work as well and take a long time to excecute for a larger amount of stocks.

Therefore the purpose of the project is to restructure the existing code in a way that improves the internal structure but doesn't change its external behavior.
We aren’t adding new functionality; but want to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read

## Results

Analysis

First I copied the starter code into the Visual Basic Editor

I reused the existing code from the starter code to 
- set up the Timer
- create the InputBox
- activate the "All Stocks Analysis" worksheet
- add the header, header row and initialize the array of tickers with the ticker values
- activate the yearValue worksheet
- and get the number of rows in the sheet to loop over

  ![image](https://user-images.githubusercontent.com/91682586/138772963-ebb74438-ba95-406b-a745-85f50c30bb10.png)


###'1a) I then created the tickerIndex that will later be used to access the arrays created in Step 1b)
    and I set it to 0. 

    tickerIndex = 0

###'1b) I created three output arrays for ticker Volumes, tickerStarterPrices, and tickerEndingPrices
    'The Datatype for ticker was set to Long, tickerStartingPrice and tickerEndingPrice to Single 
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
###'2a) I created a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To RowCount 
        tickerVolumes(tickerIndex) = 0
               
    Next i
   
    
      ###''2b) This will Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
           ### '3a) This increases the volume for the current ticker 
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value   '???
            
            ###'3b) This code checks if the current row is the first row with the selected tickerIndex and if so
            it assigns the current starting price to the tickerStartingPrice variable
                                            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
                     
        'End If
             
            ###'3c) This checks if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match the tickerIndex gets increased.
                                    
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            'End If
            
            ###'3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
                                    
            End If
            
        Next i
       
       ###'4) This code loopd through our arrays to output the Ticker, Total Daily Volume, and Return.
        
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
        Next i
        

1. 

The analysis is well described with screenshots and code (4 pt).
## Summary

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
Submission
