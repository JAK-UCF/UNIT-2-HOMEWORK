Sub StockEval()

Application.ScreenUpdating = False ' prevents screen flicker


Dim ws As Worksheet

Dim r As Long ' row number in loop

Dim CurrentTicker As String
Dim NextTicker As String

Dim OpenPrice As Double
Dim ClosePrice As Double

Dim LineVolume As Variant
Dim TotalVolume As Variant

Dim LastRow As Long ' last row of stock data to assess
    LastRow = Range("A" & Rows.Count).End(xlUp).Row

Dim PasteRow As Integer ' last row with data when entering data into summary columns


For Each ws In ThisWorkbook.Worksheets ' loop all annual worksheets
ws.Activate
    
    ' build headers for data summary columns
    Range("I1").Value = "Ticker" & vbCrLf & "Symbol"
    Range("J1").Value = "Annual" & vbCrLf & "Change:" & vbCrLf & "Value"
    Range("K1").Value = "Annual" & vbCrLf & "Change:" & vbCrLf & "Percent"
    Range("L1").Value = "Annual" & vbCrLf & "Stock" & vbCrLf & "Volume"
    ' format for uniformity
    With Range("A1:G1"): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    With Range("I1:L1"): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlTop: End With
    Rows(1).RowHeight = 50
    Range("I1:K1").ColumnWidth = 9
    Range("L1").ColumnWidth = 13
    
    
    For r = 2 To LastRow ' loop through all rows in sheet to create summary list of unique Ticker Symbols with corresponding summary data
    
    CurrentTicker = Cells(r, "A").Value
    NextTicker = Cells(r + 1, "A").Value
    LineVolume = Cells(r, "G").Value
        
        If NextTicker = CurrentTicker Then
            
            If OpenPrice = 0 Then ' get opening price on first day
                OpenPrice = Cells(r, "C").Value
            End If
            
            TotalVolume = TotalVolume + LineVolume ' add current day's volume to running total
        Else
            TotalVolume = TotalVolume + LineVolume ' add final day volume to running total
            ClosePrice = Cells(r, "F").Value ' retrieve closing price on final day
    
            PasteRow = Range("I" & Rows.Count).End(xlUp).Row + 1 ' find first open row in summary columns
            
            Cells(PasteRow, "I").Value = CurrentTicker ' paste Ticker name
            
            Cells(PasteRow, "J").Value = ClosePrice - OpenPrice ' paste data
            Cells(PasteRow, "J").NumberFormat = "#,##0.00" ' format change value
            
            If OpenPrice <> 0 Then ' error handling for stocks reporting 0 volume/opening price
                Cells(PasteRow, "K").Value = (ClosePrice - OpenPrice) / OpenPrice
            ElseIf OpenPrice = 0 And TotalVolume > 0 Then
                Cells(PasteRow, "K").Value = "0" ' return 0 if OpenPrice and TotalVolume are 0 << company reported no business
            Else
                Cells(PasteRow, "K").Value = "1" ' return 100% if Open Price is 0 and volume is not 0 << value grew '100%'
            End If
            
            Cells(PasteRow, "K").NumberFormat = "0.00%" ' format percent
            
            Cells(PasteRow, "L").Value = TotalVolume ' paste TotalVolume
            
                If (ClosePrice - OpenPrice) <= 0 Then ' conditional formatting for +/- change
                    Cells(PasteRow, "J").Interior.Color = RGB(250, 100, 100)
                    Cells(PasteRow, "K").Font.ColorIndex = 3
                Else
                    Cells(PasteRow, "J").Interior.Color = RGB(55, 150, 55)
                    Cells(PasteRow, "J").Font.Color = RGB(255, 255, 255)
                    Cells(PasteRow, "K").Font.ColorIndex = 10
                End If
    
            'RESET VARIABLE VALUES FOR NEXT LOOP
            OpenPrice = 0
            ClosePrice = 0
            TotalVolume = 0
        End If
    
    Next r
    
Next ws

Worksheets("2016").Activate
Range("A1").Select
Application.ScreenUpdating = True

MsgBox ("Multi-Year Stock Evaluation Complete!" & vbCrLf & "Please review.")

End Sub