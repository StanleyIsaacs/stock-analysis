# Analyzing Stocks Using VBA Loops
## Purpose
To provide efficient and accurate information on the performance of Green Energy stocks from the years 2017 and 2018. Client is helping his parents determine which stocks are worth investing in to provide the lowest risk and the best possible return. We'll use the data to create two charts (one for each year) that the client can access using a button inserted into the excel spreadhseet. We'd like to be able to provide the following information: the ticker, the total daily volume and the return each stock provided.

## Results

## Analysis
Before refractoring, we were able to use a previous code from our intial analysis for Steve and tweak it using nesting loops to produce a better and more efficient code. We set the initial ticker value at zero, created three different output arrays and then looped the original code to provide the final product for Steve. Below is the code used and screen shots of the run time for both years.

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

'1)Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2)Initialize an array of all tickers.

    Dim tickers(12) As String
    'Creates and array with 12 elements
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
  
'3a)Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    
'3b)Activate the data worksheet.
    Worksheets(yearValue).Activate

'3c)Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4)Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        'Do stuff with ticker
        totalVolume = 0
    
'5)Loop through rows in the data.
    Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
    '5a)Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
    End If
    
    '5b)Find the starting price for the current ticker.
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            startingPrice = Cells(j, 6).Value
        End If
        'Determines the beginning of the ticker section
        
    '5c)Find the ending price for the current ticker.
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        'Determines the end of the ticker section
        
    Next j
    
'6)Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
 Next i


Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

'Color Formatting
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
    
    If Cells(i, 3) > 0 Then
        Cells(i, 3).Interior.Color = vbGreen
        'Color the cell green
    ElseIf Cells(i, 3) < 0 Then
        Cells(i, 3).Interior.Color = vbRed
        'Color the cell red
    Else
        Cells(i, 3).Interior.Color = xlNone
        'Clear the cell color
    
    End If
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub ClearWorksheet()

Cells.Clear

End Sub

![VBA_Challenge_2017 2](https://user-images.githubusercontent.com/82114481/116839640-209a4e80-aba1-11eb-9c32-caae73062273.png)
![VBA_CHALLENGE_2018](https://user-images.githubusercontent.com/82114481/116839645-242dd580-aba1-11eb-81ae-fe8c7fef12f1.png)

## Summary

### Advantages and Disadavntages of Refactoring Code in General

Refactoring codes makes things simpler and easier for people to understand who may not be well versed in VBA coding. As mentioned above, it also makes the process more efficient. These are easier to read for everyone. It also provides a cleaner code which makes it easier to debug and better for collaboration should we choose to explore the data further. One issue that might pop up is the size of the coding files. In making results easier to present and understand, a lot of code is used. There also needs to be a test case already available for the existing code.

### Advantages and Disadvantages of the Original and Refactored VBA Script

The most obvious advantage is the speed at which the code can be ran. We were essentially able to take the run time down from around a second to less than a quarter of a second with the refactored code. The main disadvantage is the ammount of code needed and the size of the new refactored method. I had to quit Excel multiple times and reopend it after getting a '6' error. There was nothing wrong with the code, excel just needed a refresh.



