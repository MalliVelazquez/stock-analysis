# STOCK ANALYSIS

Basics on Excel might seems like something easy to work with, but, what happends if you need to look deeper into a data base?, what could we do to analyse all stocks of
a company during several years?, well, it might seems quite difficult but tuns out VBA it's a powerful tool to deliver this kind of work.
For this analysis we play a little with DAQO's data base of stock during 2017 and 2018 looking for ways to improve their capital of investment and giving Tom the oportunity 
to analyze the direction his parents business will take based on annual returns. 

##REFACTOR A CODE, PROS AND CONS

First, we'll play a little simple, we show Tom the retourns for DAQO during certain year, this might seems simple but useful. Coding for this part it's also not very elaborated:

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

You might notice this helps, but wont be useful if user does not  code and need to check another years return, nor less if he wants to know what does all this number means, so
we need to transform all this information in a more friendly user information. 

######REFACTOR CODE

We need to remember that for people that do not code it could result just OK our previous code, but they need more tools, some more reserch around the data base. For that 
porpuse we need to transform our macro in someting **easy peasy to work with**

That's why we add a new macro, able to deep dive any year they need in just a button click: 


