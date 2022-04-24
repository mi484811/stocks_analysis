# Overview of Project
### Purpose & Background
Steve reach out to create a macro to process stocks for 2017 and 2018. We will be refactoring the macro code used to process stocks for 2017 and 2018 to process more and save time the code takes to process. Steve wants his parents to use the macro to process the entire stock market over the last few years.
# Results 
### Analysis & code
I copied the code provided by Steve. Used previous code block from earlier lessons and modified them to work as indicated by Steve's outline. Several of the stocks had better return rates in 2017 than in 2018. In 2018 the US implemented a new tariff on Chinese goods. The Chinese followed the US lead and implemented a new tariff on US goods. This 2018 tariff war between the US and China led to lower return rates for individuals.

    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long, tickerStartingPrices(12) As Single, tickerEndingPrices(12) As Single

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
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then ticker = 11 Then

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

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

   

# Summary
### Advantages of Refactoring Code
The first one would be saving time. There are always new ways of writing a code that can improve the code, like an oil change to a car. Another advantage of refactoring code would be to keep it up to date since.
### Disadvantages of Refactoring Code
The time to refactor a code can be time-consuming, leading to bugs and errors. Finding solutions to the bugs and errors can be time-consuming for the team. Customers will be irritated with a product with bugs that had not previously been there. Code might not run on legacy systems.
### Advantage of Original Code
The first solution the team found to the error or outline the company wanted to accomplish. Code functions correctly in the company ecosystem.
### Disadvantage of Original Code
The original code might be outdated, using no longer used methods. Integrating code with new software can have difficulty having different standards of basic functionality.  
