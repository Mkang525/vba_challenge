Sub AlphaTest()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim WorksheetName As String
        WorksheetName = ws.Name

    Dim Ticker As String
        Ticker = " "
    Dim Total_Volume As Double
        Total_Volume = 0
    Dim Open_Price As Double
        Open_Price = 0
    Dim Close_Price As Double
        Close_Price = 0
    Dim Yearly_Change As Double
        Yearly_Change = 0
    Dim Percent_Change As Double
        Percent_Change = 0

    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    

    ws.Range("J1").Value = "Ticker"
    ws.Range("J1").Font.Bold = True
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("k1").Font.Bold = True
    ws.Range("L1").Value = "Percentage Change"
    ws.Range("L1").Font.Bold = True
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("M1").Font.Bold = True

    Open_Price = ws.Cells(2, 3).Value

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
    
         If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
         
         End If
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            
            Close_Price = ws.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price
    
            If Open_Price <> 0 Then
            Percent_Change = (Yearly_Change / Open_Price) * 100
            
         End If
        
    ws.Range("J" & Summary_Table_Row).Value = Ticker
    ws.Range("K" & Summary_Table_Row).Value = Yearly_Change

    If (Yearly_Change > 0) Then
        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf (Yearly_Change <= 0) Then
        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    End If

     ws.Range("L" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
     ws.Range("M" & Summary_Table_Row).Value = Total_Volume

     Summary_Table_Row = Summary_Table_Row + 1
     Yearly_Change = 0
     Close_Price = 0
     Open_Price = ws.Cells(i + 1, 3).Value

End If
Next i
Next ws

End Sub

