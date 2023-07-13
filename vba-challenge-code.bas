Attribute VB_Name = "Module1"

Sub vba_challenge()
'define variables
Dim ticker As String
Dim vol As Double
vol = 0
Dim YrOpen As Double
YrOpen = 0
Dim YrClose As Double
YrClose = 0
Dim YrChange As Double
Change = 0
Dim PercentChange As Double
PercentChange = 0
    
'Create new columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"



'loop source: https://excelchamps.com/vba/usedrange/#:~:text=1%20Write%20a%20Code%20with%20UsedRange.%20Use%20the,code%20shows%20a%20message%20box%20with%20the%20
'source: https://stackoverflow.com/questions/12997645/what-operator-is-in-vba
'loop so that for each new string new row is created
For i = 2 To ActiveSheet.UsedRange.Rows.Count
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
    vol = Cells(i, 7).Value
    YrOpen = Cells(i, 3).Value
    YrClose = Cells(i, 6).Value
    YrChange = YrOpen - YrClose
    PercentChange = YrChange / YrOpen
    Dim Row1 As Integer
    Row1 = 2
    Cells(Row1, 9).Value = ticker
    Cells(Row1, 10).Value = YearChange
    Cells(Row1, 11).Value = PercentChange
    Cells(Row1, 12).Value = vol
    Row1 = Row1 + 1
  End If
    
Next i

'I am stuck because it runs the program but it will only stay in the first row
'I am not sure how to make this work for multiple worksheets

End Sub
