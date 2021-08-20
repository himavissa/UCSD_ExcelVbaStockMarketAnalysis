# UCSD_ExcelVbaStockMarketAnalysis
"Define Module and assign this macros on the main worksheet

Sub GetSummary()
"Declare variable

Dim Ticker As String
Dim Stryear As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearChange As Double
Dim PerChange As Double
Dim TickerVol As Double
Dim i As Long
Dim RowCnt As Integer
Dim FirstTime As Integer
Dim GrPerIncValue As Double
Dim GrPerIncTicker As String
Dim GrPerDecValue As Double
Dim GrPerDecTicker As String
Dim GrTotalVol As Double
Dim GrtotalTicker As String


'Intialize 1st row index for Main loop
i = 2

' initialize the row count for summary table

RowCnt = 2


OpenPrice = Cells(2, 3).Value

'Assigning header

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "OpenPrice"
Cells(1, 11).Value = "ClosePrice"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "Percent Change"
Cells(1, 14).Value = "Total Stock Volume"

GrPerIncValue = 0
GrPerDecValue = 0
GrTotalVol = 0
GrPerIncTicker = ""
GrPerDecTicker = ""
GrtotalTicker = ""



While Cells(i, 1) <> ""
Ticker = Cells(i, 1).Value
Stryear = Left(Cells(i, 2).Value, 4)

'Initialize the Price

'ClosePrice = 0
TickerVol = 0
FirstTime = 0


'Checking condition with while if cells are not empty and for ticker values

While Cells(i, 1) <> "" And Ticker = Cells(i, 1).Value And Stryear = Left(Cells(i, 2).Value, 4)

If FirstTime = 0 Then

OpenPrice = Cells(i, 3).Value
FirstTime = 1

End If

ClosePrice = Cells(i, 6).Value
TickerVol = TickerVol + Cells(i, 7).Value

i = i + 1

Wend

'Calculate Yearly Change
yearlyChange = (ClosePrice - OpenPrice)

If OpenPrice = 0 Then
ChangePer = 100
Else
ChangePer = (100 * yearlyChange) / OpenPrice

End If

Range("I" & RowCnt).Value = Ticker
Range("J" & RowCnt).Value = OpenPrice
Range("K" & RowCnt).Value = ClosePrice
Range("L" & RowCnt).Value = yearlyChange
Range("M" & RowCnt).Value = ChangePer
Range("N" & RowCnt).Value = TickerVol

' To check for the max greatest %

If ChangePer > GrPerIncValue Then
GrPerIncValue = ChangePer
GrPerIncTicker = Ticker
End If

If ChangePer < GrPerDecValue Then
GrPerDecValue = ChangePer
GrPerDecTicker = Ticker
End If

If TickerVol > GrTotalVol Then
GrTotalVol = TickerVol
GrtotalTicker = Ticker
End If


'Color Index for Change %
If ChangePer > 0 Then
Range("L" & RowCnt & ":L" & RowCnt).Interior.ColorIndex = 4

End If
If ChangePer < 0 Then
Range("L" & RowCnt & ":L" & RowCnt).Interior.ColorIndex = 3

End If


RowCnt = RowCnt + 1

Wend

Range("P3").Value = "Greatest % Increase"
Range("P4").Value = "Greatest % Decrease"
Range("P5").Value = "Greatest Total Volume"

Range("Q2").Value = "Ticker"
Range("R2").Value = "Value"

Range("Q3").Value = GrPerIncTicker
Range("R3").Value = GrPerIncValue
Range("Q4").Value = GrPerDecTicker
Range("R4").Value = GrPerDecValue
Range("Q5").Value = GrtotalTicker
Range("R5").Value = GrTotalVol







End Sub

