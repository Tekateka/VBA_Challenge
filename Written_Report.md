## Module 2 - Stock Analysis 

## Overview of Project
 Steve would like to analyze the stocks data using VBA. The stocks market dataset contains two years data for 2017 and 2018. Steve also would like to access the performance of the stocks. With the original script the run may take a long time to excute, therefore the codes need to be refractored to excuite faster and optimize the existing script.   

### Purpose
The purpose of the project was to compare the stock performance between 2017 and 2018. And to access and compare the execution times of the original script and refractored script.  
 
## Analysis and outcomes 
The data were analysed on excel using VBA and tables were created for visualization.  

 '1a) Create a ticker Index
     Dim tickerIndex As Integer
     tickerIndex = 0
     
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     
    For i = 0 To 11
        tickerStartingPrice(i) = 0
        tickerEndingPrice(i) = 0
        
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
    '3a Increase volume for current ticker
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b Check if the current row is the first row with the selected tickerIndex.
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
    
    End If
    
    '3c check if the current row is the last row with the selected ticker
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
    End If
    
    'Increase the tickerIndex.
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
    
    End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
        
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1

        
    Next i
    

## Results  
  #  Outputs for year 2017 and 2018 
VBA analysis for stock performances and excution times for 2017 and 2018 are attached. 

The analysis done for stock performances for two years were similar as shown in module 2. 
Based on the analysis in year 2017 except ticker, TERP; the rest stocks showed positive returns whereas in 2018 only two stocks (ENPH and RUN) showed positive returns. Therefore, year 2017 stocks had higher positive returs as compared to 2018. 

The excution time for refactored Vs original script shows that the refactored scripts were much quicker that the original scripts. 

## Summary 

# Advantages and disadvantages of refractored code

  # Advantages 
- We remove excess or redundant codes while refractoring so it's easier to understand and modify the codes
- Writing codes are quicker and executing faster
- Improves codes quality and the design of existing code, and can help with debugging 
- Use to fix bugs, clean, and organize codes. 
- Refractoring is always an option 

  # Disadvantages 
- can break working code that is difficult to understand and poorly structured 
- It may introduce bugs and risky in the view of management 
- The process requires skills and discipline
 

# Advantages and disadvantages of original and refractored VBA script
   # Advantages 
 The codes in VBA analysis runs quickly and it was simple to read compared to the original script. Code refractoring is a way of restructuring and optimizing existing code without changing its external behavior. Refractoring makes to run the codes quicker, easy to read and undestand than the original script. Also enable the user to review and excuite the whole data easily than original script.  

-As for disadvanatage, using the original script the VBA only allow one year data to analyze and review. One of the disadvatages of refaractoring is while refractoring we may miss codes and break the codes; in this situation the process can be risky and time consuming. 