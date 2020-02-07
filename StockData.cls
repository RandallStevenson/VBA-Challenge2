VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data():

Dim ws As Worksheet
For Each ws In Sheets
ws.Activate

'Set stock data variable dimensions
Dim i As Long
Dim Stock_No As Long
Dim Ticker As String
Dim Change As Double
Dim PercChg As Double
Dim Open_Price As Double
'Since volume is greater than 32767, set dimension as Long
Dim Volume As Double
Dim GPI As Double
Dim GPD As Double
Dim GVol As Double

'Put Tiltes on Cells Columns/Rows
Cells(1, 10) = "Ticker"
Cells(1, 11) = "Yearly Change"
Cells(1, 12) = "Percent Change"
Cells(1, 13) = "Total Stock Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest Percent Increase"
Cells(3, 15) = "Greatest Percent Decrease"
Cells(4, 15) = "Greatest Volume"

'Input Initial Values
Ticker = Cells(2, 1)
Open_Price = Cells(2, 3)
Volume = 0
GPI = 0
GPD = 0
GVol = 0
Stock_No = 2

'Put Initial Values on Spreadsheet
Cells(2, 16) = Ticker
Cells(4, 16) = Ticker
Cells(4, 16) = Ticker
Cells(2, 17) = 0
Cells(3, 17) = 0
Cells(4, 17) = 0

'Get stock data
i = 2
Do While (Cells(i, 1) <> "")
    'Increment volume of Current Stock
    Volume = Volume + Cells(i, 7)
    'If next cell changes ticker then output information
    If Cells(i, 1) <> Cells(i + 1, 1) Then
        Cells(Stock_No, 10) = Ticker
        Change = Cells(i, 6) - Open_Price
        If Open_Price > 0 Then
                PercChg = Change / Open_Price
            Else
                PercChg = 0
        End If
        Cells(Stock_No, 11) = Change
        Cells(Stock_No, 12) = FormatPercent(PercChg, 2, True)
        Cells(Stock_No, 13) = Volume
        If Change < 0 Then
                Cells(Stock_No, 11).Interior.ColorIndex = 3
            Else
                Cells(Stock_No, 11).Interior.ColorIndex = 4
        End If
        If PercChg > GPI Then
                GPI = PercChg
                Cells(2, 16) = Ticker
            ElseIf PercChg < GPD Then
                GPD = PercChg
                Cells(3, 16) = Ticker
        End If
        If Volume > GVol Then
            GVol = Volume
            Cells(4, 16) = Ticker
        End If
        If Cells(i + 1, 1) <> "" Then
            Ticker = Cells(i + 1, 1)
            Open_Price = Cells(i + 1, 3)
            Volume = 0
        End If
        Stock_No = Stock_No + 1
    End If
i = i + 1
Loop


Cells(2, 17) = FormatPercent(GPI, 2, True)
Cells(3, 17) = FormatPercent(GPD, 2, True)
Cells(4, 17) = GVol


Next ws

End Sub

