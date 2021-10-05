# stocks-analysis
Overview of Project
Our client tasked us with creating a tool to efficiently and accurately analyze stock data. He initially gave us 12 stocks to begin working with, but he would like to be able to use the tool for any stock that he or his clients (parents) are interested in.
Results
The results of the stock performance are shown below.
  Images showing stock results are loaded as PNG files.
            
Hopefully they bought in 2018! Or else they might have purchased at inflated prices. Companies that produce solar panels or equipment for solar panels are in a unique space. They can be considered energy companies or technology companies. Since it is a relatively new space these companies can see large swings in stock price that are outside of the companies control. It is not surprising to see prices become inflated as investors try to pick out who will be the dominant players in solar power in the next 5 to 10 years. With that being said, stock price alone is not enough to determine which companies stock to buy. 
Our original code produced a run time shown below.
     
Given a small group of stocks this amount of time is not a big deal. But once we refactored the code we were able to reduce the run time which I will show here.
  Images showing the time saved are loaded as PNG files.

We can see that the refactored code is quicker and it is more versatile. The original code required that we provide the ticker symbols. Which is a very manual task and not very much fun. The refactored code is not only quicker but we do not need to provide the ticker symbols. We just run the code. 
Summary
The advantages of the original code is that we know that it works. I cannot see any other obvious advantages. The advantages of the refactored code are quite obvious. It is faster and we do not need to hardcode anything. Another disadvantage of both sets of code is that we need to know how the data will be presented prior to running the code. If the symbols were not grouped or if there were extra columns we may need to adjust our code. That is a disclaimer I would give the client so that he is aware in case he gets a very different output than he is expecting.



ORIGINAL CODE

Sub AllStocksAnalysis()
'1)Format the output sheet on the "All Stocks Analysis" worksheet.
    Dim sartTime As Single
    Dim endTime As Single
    
    Worksheets("All Stocks Analysis").Activate
        yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
        Range("A1").Value = "All Stocks" + yearValue

        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
    
'2)Initialize an array of all tickers.
    
    Dim tickers(11) As String
    
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

'3)Prepare for the analysis of tickers.
    '3a)Initialize variables for the starting price and ending price.
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b)Activate the data worksheet.
    
    Sheets(yearValue).Activate
    
    '3c)Find the number of rows to loop over.
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4)Loop through the tickers.
    
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
'5)Loop through rows in the data.
        Sheets(yearValue).Activate
        For j = 2 To RowCount
        
    '5a)Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
        
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
    '5b)Find the starting price for the current ticker.
    
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            startingPrice = Cells(j, 6).Value

        End If
        
    '5c)Find the ending price for the current ticker.
    
     If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            endingPrice = Cells(j, 6).Value

        End If
        Next j
        
'6)Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
Next i
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

        Worksheets("All Stocks Analysis").Activate
    Range("A3:c3").Font.Bold = True
    Range("A3:c3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:c3").Borders(xlEdgeBottom).Color = RGB(0, 0, 255)
    Range("A3:C3").Font.Color = RGB(0, 0, 255)
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:c15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
            
        Else
            Cells(i, 3).Interior.Color = xlNone
            
        End If
        
            
    
    Next i
    
    

End Sub



Sub YearValueAnalysis()
yearValue = InputBox("What year would you like to run the analysis on?")
End Sub

REFACTORED CODE

Sub AllStocksAnalysis()
'1)Format the output sheet on the "All Stocks Analysis" worksheet.
Dim sartTime As Single
Dim endTime As Single

Worksheets("All Stocks Analysis").Activate
yearValue = InputBox("What year would you like to run the analysis on?")
startTime = Timer
Range("A1").Value = "All Stocks" + yearValue

'Create a header row
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

'needed variables
Dim startingPrice As Single
Dim endingPrice As Single
Dim tickerVolumes As Long

'3b)Activate the data worksheet.

Sheets(yearValue).Activate

'3c)Find the number of rows to loop over.

RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
'4)Loop through the tickers.
Dim tickerIndex As Integer
tickerIndex = 0
Dim tickerCounter As Integer
tickerCounter = 0

ticker = Cells(2, 1)
totalVolume = 0
startingPrice = 0
endingPrice = 0
   
'begin loop
For i = 2 To RowCount + 1
tickerIndex = tickerIndex + 1

' STARTING Price
If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
startingPrice = Cells(i, 6).Value
End If


'ticker = Cells(i, 1)
If Cells(i, 1) <> ticker Then
tickerIndex = 0

endingPrice = Cells(i - 1, 6).Value

'paste output
Worksheets("All Stocks Analysis").Activate
Cells(4 + tickerCounter, 1).Value = ticker
Cells(4 + tickerCounter, 2).Value = totalVolume
Cells(4 + tickerCounter, 3).Value = (endingPrice / startingPrice) - 1

'reset for next ticker
Sheets(yearValue).Activate
ticker = Cells(i, 1)
totalVolume = 0
startingPrice = 0
endingPrice = 0

tickerCounter = tickerCounter + 1
startingPrice = Cells(i, 6).Value

End If

' TOTAL VOLUME
'If Cells(i, 1).Value = ticker Then
totalVolume = totalVolume + Cells(i, 8).Value
'End If

Next i

'function timer
endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

'output formatting
Worksheets("All Stocks Analysis").Activate
Range("A3:c3").Font.Bold = True
Range("A3:c3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A3:c3").Borders(xlEdgeBottom).Color = RGB(0, 0, 255)
Range("A3:C3").Font.Color = RGB(0, 0, 255)
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:c15").NumberFormat = "0.0%"
Columns("B").AutoFit

dataRowStart = 4
dataRowEnd = 15
For i = dataRowStart To dataRowEnd

If Cells(i, 3) > 0 Then
Cells(i, 3).Interior.Color = vbGreen

ElseIf Cells(i, 3) < 0 Then
Cells(i, 3).Interior.Color = vbRed

Else
Cells(i, 3).Interior.Color = xlNone

End If



Next i



End Sub




