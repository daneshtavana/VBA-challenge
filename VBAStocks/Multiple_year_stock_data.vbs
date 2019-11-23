Attribute VB_Name = "Module1"
Sub Sheet_All()
  ' This subroutine will update all sheets
  Call Sheet_2014
  Call Sheet_2015
  Call Sheet_2016
End Sub

Sub Sheet_2014()
  ' This subroutine sets the sheet variable,
  ' then calls and passes it to TickerList
  Set WS = Sheets("2014")
  Call TickerList(WS)
End Sub

Sub Sheet_2015()
  ' This subroutine sets the sheet variable,
  ' then calls and passes it to TickerList
  Set WS = Sheets("2015")
  Call TickerList(WS)
End Sub

Sub Sheet_2016()
  ' This subroutine sets the sheet variable,
  ' then calls and passes it to TickerList
  Set WS = Sheets("2016")
  Call TickerList(WS)
End Sub

Sub TickerList(WS)
  ' Create variables an initialize
  Dim Symbol As String
  Dim Volume As Double
  Dim BeginingClosePrice As Currency
  Dim EndingClosePrice As Currency
  Dim YearlyChange As Currency
  Dim PecentChange As Double
  Dim OutRowNum As Integer
  OutRowNum = 2
  
  ' Initialize WS.Cells, just a precaution to avoid divide by zero
  Symbol = JUNK
  BeginingClosePrice = 10
  EndingClosePrice = 10
  YearlyChange = 10
  PercentChange = 10
  Volume = 0
  
  ' Count the number of input rows to loop and format appropriate column
  lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
   
  ' Loop through each row, use lastrow variable to end Loop'
  ' When TickerSymbol change is detected log new information)
  For i = 2 To lastrow + 1
    If Symbol <> WS.Cells(i, 1).Value Then
       Symbol = WS.Cells(i, 1).Value
       YearlyChange = EndingClosePrice - BeginingClosePrice
       If BeginingClosePrice = 0 Then
           MsgBox ("Error in Row " & OutRowNum & " Beginning Close Price = zero")
           PercentChange = 0
       Else
           PercentChange = YearlyChange / BeginingClosePrice
       End If
       ' Ticker Symbol changes, so log the row values
       WS.Cells(OutRowNum, 9).Value = Symbol
       WS.Cells(OutRowNum - 1, 10).Value = YearlyChange
       ' Color code Yearly Change, Column 10, when values are less than zero
       If YearlyChange < 0 Then
           WS.Cells(OutRowNum - 1, 10).Interior.ColorIndex = 3
       Else
           WS.Cells(OutRowNum - 1, 10).Interior.ColorIndex = 4
       End If
       WS.Cells(OutRowNum - 1, 11).Value = PercentChange
       ' Format Percent Change, Column 11
       WS.Cells(OutRowNum - 1, 11).NumberFormat = "0.00%"
       WS.Cells(OutRowNum - 1, 12).Value = Volume
       ' WS.Cells were inserted but later commented only to assist with debug
       ' WS.Cells(OutRowNum - 1, 13).Value = BeginingClosePrice
       ' WS.Cells(OutRowNum - 1, 14).Value = EndingClosePrice
       ' Point to new Log output row
       OutRowNum = OutRowNum + 1
       ' Since Ticker Symbol changed, initialize closing date and volume
       BeginingClosePrice = WS.Cells(i, 6).Value
       Volume = 0
    Else
       EndingClosePrice = WS.Cells(i, 6).Value
       Volume = Volume + WS.Cells(i, 7).Value
    End If
    Next i
    
  ' Create Output Report Header Rows
   WS.Cells(1, 9).Value = "Ticker"
   WS.Cells(1, 10).Value = "Yearly Change"
   WS.Cells(1, 11).Value = "Percent Change"
   WS.Cells(1, 12).Value = "Volume"
   ' WS.Cells were inserted but later commented only to assist with debug
   ' WS.Cells(1, 13).Value = "Beginning Price"
   ' WS.Cells(1, 14).Value = "Ending Price"
   
  ' This is another method to Format columns
  ' Range("K2:K" & lastrow).NumberFormat = "0.00%"
  
  End Sub
