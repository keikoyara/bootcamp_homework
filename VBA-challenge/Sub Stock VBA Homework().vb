Sub Stock_VBA_Homework()

' loop through all sheets
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

'find last row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
 
 
'Add heading for new columns
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"

'Create Variables
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Ticker_Name As String
Dim Percent_Change As Double
Dim Volume As Double
Volume = 0
Dim Row As Double
Row = 2
Dim Column As Integer
Column = 1
Dim i As Long

'initial open price
Open_Price = Cells(2, Column + 2).Value

'loop through all symbols
For i = 2 To LastRow
    
    'run through first symbol, is it the same ticker symbol? If not,
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        
        'set Ticker Name
        Ticker_Name = Cells(i, Column).Value
        Cells(Row, Column + 8).Value = Ticker_Name
        
        'set close price
        Close_Price = Cells(i, Column + 5).Value
        
        'add yearly change
        Yearly_Change = Close_Price - Open_Price
        Cells(Row, Column + 9).Value = Yearly_Change
        
        'add percent change
        If (Open_Price = 0 And Close_Price = 0) Then
            Percent_Change = 0
            
        ElseIf (Open_Price = 0 And Close_Price <> 0) Then
            Percent_Change = 1
            
        Else
            Percent_Change = Yearly_Change / Open_Price
            Cells(Row, Column + 10).Value = Percent_Change
            Cells(Row, Column + 10).NumberFormat = "0.00%"
            
        End If
        
        'add total stock volume
        Volume = Volume + Cells(i, Column + 6).Value
       Cells(Row, Column + 11).Value = Volume
        
        'add one to row
        Row = Row + 1
        
        'reset the open price
        Open_Price = Cells(i + 1, Column + 2)
        
        'reset the volume total
        Volume = 0
        
        'if cells are the same ticker
        Else
            Volume = Volume + Cells(i, Column + 6).Value
       
    End If
    Next i

'find last row of yearly change
 YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
 
 'set color of cells
 For j = 2 To YCLastRow
    If Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0 Then
        Cells(j, Column + 9).Interior.ColorIndex = 10
    ElseIf Cells(j, Column + 9).Value < 0 Then
        Cells(j, Column + 9).Interior.ColorIndex = 3
        
    End If
    Next j
    

Next WS

End Sub
