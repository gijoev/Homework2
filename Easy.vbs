Sub easyWorksheet()

Dim WS_Count As Integer
Dim I As Integer
        WS_Count = ActiveWorkbook.Worksheets.Count
        MsgBox WS_Count
        
Dim LastRow As Long
Dim ws As Worksheet
   For Each ws In ActiveWorkbook.Worksheets
   ws.Activate
       
       Dim ticker As Long
       Dim vol As Long
       Dim total As Double
           vol = 2
           LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
           Cells(1, 10) = "Total Stock Volume"
           Cells(1, 9) = "Ticker"
               For ticker = 2 To LastRow
                    If Cells(ticker, 1) <> Cells(ticker + 1, 1) Then
                         Cells(vol, 9) = Cells(ticker, 1)
                         Cells(vol, 10) = total + Cells(ticker, 7)
                         vol = vol + 1
                         total = 0
                    Else
                    total = total + Cells(ticker, 7)
                   End If
               Next
          MsgBox ws.Name
   Next
 End Sub