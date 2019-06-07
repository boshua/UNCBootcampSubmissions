Attribute VB_Name = "Module1"
Sub Workbook_testing():

For Each ws In Worksheets

'Find Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create Column Headers

    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Total Stock Volume"
    ws.Cells(1, "K").Value = "Yearly Change"
    ws.Cells(1, "L").Value = "Percent Change"

'Create Variables

    Dim Volume As Double
    Volume = 0
    Dim Row As Double
    Row = 2
    Dim Column As Integer
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Ticker_Name As String
    Dim i As Long

'Set initial open price
    Opening_Price = ws.Cells(2, 3).Value

'Loop through all data
    For i = 2 To LastRow

'Check to see if ticker symbol has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
'Set ticker name
            Ticker_Name = ws.Cells(i, 1).Value
            
'Apply ticker name to sub table
            ws.Cells(Row, 9).Value = Ticker_Name

'Set Closing Price
            Closing_Price = ws.Cells(i, 6).Value
            
'Add Yearly Change
            Yearly_Change = Closing_Price - Opening_Price

'Apply Yearly Change to Summary Table
            ws.Cells(Row, 11).Value = Yearly_Change

'Determine Percent Change
            If (Opening_Price = 0 And Closing_Price = 0) Then
                Percent_Change = 0
            
            ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                Percent_Chnange = 1
            
            Else
                Percent_Change = Yearly_Change / Opening_Price
 
 'Apply Percent Change to Summary Table
                ws.Cells(Row, 12).Value = Percent_Change
                ws.Cells(Row, 12).NumberFormat = "0.00%"
            End If
            
'Find Total Volume
            Volume = Volume + ws.Cells(i, 7).Value
            
'IApply Volume Data into Summary Table
            ws.Cells(Row, 10).Value = Volume

'Move Row to Next Slot in Summary Table
            Row = Row + 1

'Reset the Opening Price
            Opening_Price = ws.Cells(i + 1, 3)

'Reset Volume for Next Stock
            Volume = 0
'If Ticker Cells are the Same
        Else
            Volume = Volume + ws.Cells(i, 7).Value
              
        End If
    
    Next i
    
'Find Last Row of Yearly Change
    YCLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
'Set the Cell Colors
    For j = 2 To YCLastRow
        If (ws.Cells(j, 11).Value > 0 Or ws.Cells(j, 11).Value = 0) Then
            ws.Cells(j, 11).Interior.ColorIndex = 10
        
        ElseIf ws.Cells(j, 11).Value < 0 Then
            ws.Cells(j, 11).Interior.ColorIndex = 3
        
        End If
        
    Next j
          

        
        
    
Next ws

End Sub






