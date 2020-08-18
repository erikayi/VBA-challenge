Attribute VB_Name = "Module2"
Sub stock()

Dim Ticker As String
Dim Volume As Double
Volume = 0

Dim Summary_Table_Row As Integer
Dim Year_Open As Double
Dim Year_close As Double
'Dim Year_High As Double
'Dim Year_Low As Double

'For Each ws In Worksheets
'MsgBox (ws.Name)
'Set ws = Worksheet("ws")

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Total_Stock_Volume"
Cells(1, 12).Value = "Yearly_Percentage"

Summary_Table_Row = 2

Dim LastRow As Integer
Dim i As Integer
    
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
            If Year_Open = 0 Then
            
            Year_Open = Cells(i, 3).Value
            
            End If
            
            Next i
            
            If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Year_close = Cells(i, 6).Value
                Yearly_Percentage = Cells(i, 5).Value
                Yearly_Change = Year_close - Year_Open
                Yearly_Percentage = (Year_close - Year_Open) * 100
            
            Ticker = Cells(i, 1).Value
            
            Volume = Volume + Cells(i, 7).Value
            
            Range("J" And Summary_Table_Row).Value = Yearly_Change
                           
            Range("I" And Summary_Table_Row).Value = Ticker
            Range("K" And Summary_Table_Row).Value = Yearly_Percentage
            Range("L" And Summary_Table_Row).Value = Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Volume = 0
            
   End If
          

End Sub
