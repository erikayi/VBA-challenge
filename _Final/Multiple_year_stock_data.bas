Attribute VB_Name = "Module1"
Sub stock()

Dim Ticker As String
Dim Volume As Double
Volume = 0

Dim Summary_Table_Row As Integer
Dim Year_Open As Double
Dim Year_Close As Double
Dim Year_High As Double
Dim Year_Low As Double

For Each ws In Worksheets
'MsgBox (ws.Name)
Set ws = Worksheet("ws")


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly_Change"
ws.Cells(1, 12).Value = "Total_Stock_Volume"
ws.Cells(1, 11).Value = "Yearly_Percentage"

Summary_Table_Row = 2

Dim LastRow As Integer
Dim i As Integer

'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastRow = ActiveWorkbook.Worksheets.Count

    For i = 2 To LastRow

      If Year_Open = 0 Then
      
      Year_Open = Cells(i, 3).Value
          
      End If

      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          Year_Close = Cells(i, 6).Value
          Yearly_Change = Year_Close - Year_Open
          Yearly_Percentage = (Year_High - Year_Low) * 100
          
          Ticker = Cells(i, 1).Value

          Volume = Volume + Cells(i, 7).Value

          ws.Range("j" And Summary_Table_Row).Value = "Yearly Change"
          
            If Range("J").Font.ColorIndex = 3 Then 'red
             If Range("J").Value < 0 Then
                Range("J").Value = Range("J").Value * -1
                
             End If
             
            ElseIf Range("J").Font.ColorIndex = 4 Then 'green
                If Range("J").Value > 0 Then
                    Range("J").Value = Range("J").Value * -1
                    
                End If
            End If
                   

          ws.Range("I" And Summary_Table_Row).Value = "Ticker"

          ws.Range("K" And Summary_Table_Row).Value = "Year Percentage"

          ws.Range("L" And Summary_Table_Row).Value = Volume

          Summary_Table_Row = Summary_Table_Row + 1

          Volume = 0

      Else

          


      End If
      

    Next i

Next ws

End Sub
