# STOCK ANALYSIS

## WHAT TO ANALYSE FOR?

Basics on Excel might seems like something easy to work with, but, what happends if you need to look deeper into a data base?, what could we do to analyse all stocks of
a company during several years?, well, it might seems quite difficult but tuns out VBA it's a powerful tool to deliver this kind of work.
For this analysis we play a little with DAQO's data base of stock during 2017 and 2018 looking for ways to improve their capital of investment and giving Tom the oportunity 
to analyze the direction of his parents business will take based on annual returns. 

## DQ's ANNUAL RETURN

Tom's parents woul like to analyse DAQO's annual return in order to take decisions about their investments. We create an easy way to compare and analyse this desicion. 
with our code we can show them annual returns summarized and well ordered. 

As showed bellow, Tom's parents has a lot to think about DQ. 2017 vs 2018 results does not seems promissing. It looks like company has to work hard in order to change this 
red numbers. Maybe another strategy or a deep thinking about the future might help: 

![2017](https://user-images.githubusercontent.com/96633294/149430347-f806cb10-776f-42ce-9ef8-b2c7c0173f86.png)

![2018](https://user-images.githubusercontent.com/96633294/149431212-87fb3079-a72f-4072-a881-658324b1ac2d.png)


## REFACTOR A CODE, PROS AND CONS

First, we'll play a little simple, we show Tom DAQO's retourn during certain year, this might seems simple but useful. Coding for this part it's also not very elaborated:

```
    Worksheets("2018").Activate
    
    rowStart = 2
    'DELETE rowEnd = 3013
    'Row code taken https://www.youtube.com/watch?v=WwJzVDY6lf8
    rowEnd = ThisWorkbook.Sheets("2018").Cells(Rows.Count, 1).End(xlUp).Row
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    For i = rowStart To rowEnd
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
        totalVolume = totalVolume + Cells(i, 8).Value
        End If

        'Calculate startin/ending prices
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        startingPrice = Cells(i, 6).Value
        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        endingPrice = Cells(i, 6).Value
        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    
  ```

You might notice this helps, but wont be useful if user does not code and need to check another years return, nor less if he wants to know what does all this numbers means, so
we need to transform all this information in a more friendly user information. 

###### REFACTOR CODE

We need to remember that for people that do not code it could result just OK our previous code, but they need more tools, some more reserch around the data base, but keep it
simple for them. For that porpuse we need to transform our macro in someting more **easy peasy to work with**

That's why we add a new macro, able to deep dive any year they need in just a button click and showing with colors what does all that numbers mean: 

```
''2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
    tickerIndex = tickers(i)
    tickerVolumes = 0
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    Sheets(yearvalue).Activate
    
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = tickerIndex Then
        tickerVolumes = tickerVolumes + Cells(j, 8).Value
    End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
        startingPrice = Cells(j, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
        endingPrice = Cells(j, 6).Value
        End If
    
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = tickerVolumes
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    
    If Cells(4 + i, 3) > 0 Then
    'Colored green 
        Cells(4 + i, 3).Interior.Color = vbGreen
        
    ElseIf Cells(4 + i, 3) < 0 Then
    'Colored red
        Cells(4 + i, 3).Interior.Color = vbRed
        
    End If
```
You can notice this "new code" looks pretty similar as the one before but it takes all worksheets and run it at the same time. Pretty simple as it sounds, user can just click a
button and all their information we'll be showed. 

Only thing left is to take decitions, to analyse the direction of the company and invest in the best option. 

**great power comes great responsibility**

