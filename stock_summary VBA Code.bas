Attribute VB_Name = "Module2"
 Sub stock_summary():
 For Each ws In Worksheets
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Single
    Dim Total_Volume As LongLong
    Dim Summary_Table_Row As Integer
    Dim LastRow As Long
    Dim Volume As Long
    Dim Yearly_Open As Double
    Dim Yearly_Close As Double
    
    Summary_Table_Row = 2

    Volume = 0
    Total_Volume = 0

    Yearly_Open = ws.Cells(2, 3).Value
    Yearly_Close = 0
    Yearly_Change = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("L1").Value = "Total Stock Volume"
    
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                Ticker = ws.Cells(i, 1).Value

                Volume = ws.Cells(i, 7).Value
                Total_Volume = Total_Volume + Volume
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume

                Yearly_Close = ws.Cells(i, 6).Value
                
                Yearly_Change = Yearly_Close - Yearly_Open
                
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                If Yearly_Open <> 0 Then
                    Percent_Change = (Yearly_Change / Yearly_Open)
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                Else
                    ws.Range("K" & Summary_Table_Row).Value = 0
                End If

                    If Yearly_Change > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If

                Yearly_Open = ws.Cells(i + 1, 3).Value

                Summary_Table_Row = Summary_Table_Row + 1

                Total_Volume = 0

            Else
                Volume = ws.Cells(i, 7).Value
                Total_Volume = Total_Volume + Volume
            End If
        
        Next i
    
    Next ws
 
 End Sub
 
