Sub moderate()

Dim WS_Count As Integer
Dim I As Integer
         WS_Count = ActiveWorkbook.Worksheets.Count
         MsgBox WS_Count
Dim NumRows As Long
Dim VolumeCount As Double
Dim yearclose As Double
Dim yearopen As Double
Dim yeardiff As Double
Dim SolutionRow As Long
Dim LastRow As Long
Dim ws As Worksheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'For Each ws In ActiveWorkbook.Worksheets
    'ws.Activate
For Row = 2 To LastRow
        'These are the middle rows for each stock, take and tally the volume count
       If Cells(Row - 1, 1) = Cells(Row, 1) And Cells(Row + 1, 1) = Cells(Row, 1) Then
           VolumeCount = VolumeCount + Cells(Row, 7)
           Cells(SolutionRow, 10) = VolumeCount
    'Check to replace a yearopen of zero
           If yearopen = 0 Then
           yearopen = Cells(Row, 3)
           End If
       
       'This is the last row for each stock, take the year close value
       ElseIf Cells(Row - 1, 1) = Cells(Row, 1) And Cells(Row + 1, 1) <> Cells(Row, 1) Then
           yearclose = Cells(Row, 6)
           yeardiff = yearclose - yearopen
           Cells(SolutionRow, 11) = yeardiff
           'Cells(SolutionRow, 11).NumberFormat = "0"
       'Apply conditional formatting for positive/negative absolute changes
               If yeardiff > 0 Then
                   Cells(SolutionRow, 11).Font.Color = RGB(0, 130, 53)
                   Cells(SolutionRow, 11).Interior.Color = RGB(226, 239, 218)
               ElseIf yeardiff < 0 Then
                   Cells(SolutionRow, 11).Font.Color = RGB(192, 0, 0)
                   Cells(SolutionRow, 11).Interior.Color = RGB(255, 215, 214)
               End If
               
       'Control for a yearopen of zero (if a stock had a year’s worth of prices all 0 like PLNT)
           If yearopen <> 0 Then
               yearpercent = yeardiff / yearopen
               Cells(SolutionRow, 12) = yearpercent
               Cells(SolutionRow, 12).NumberFormat = "0.000%"
               Else
               yearpercent = 0
               Cells(SolutionRow, 12) = yearpercent
           End If
           
       'This is the first instance/row of each stock, get name, volume and opening price
       ElseIf Cells(Row - 1, 1) <> Cells(Row, 1) Then
           SolutionRow = SolutionRow + 1
           StockName = Cells(Row, 1)
           Cells(SolutionRow, 9) = StockName
           VolumeCount = Cells(Row, 7)
           Cells(SolutionRow, 10) = VolumeCount
           yearopen = Cells(Row, 3)
       End If
  Next

  End Sub

