Attribute VB_Name = "Module1"

Sub test45()

'Define variables
Dim oldSymbol As String
Dim newSymbol As String
Dim totalVol As Double
Dim nextVol As Double
Dim printTickerRow As Integer
printTickerRow = 2

'Count the number of rows of data
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set initial Vol for ticker A
totalVol = Cells(2, 7)

For i = 2 To lastRow
       
    oldSymbol = Cells(i, 1).Value
    
    newSymbol = Cells(i + 1, 1).Value
        
    nextVol = Cells(i + 1, 7).Value
        
    ' New ticker symbol different
    If newSymbol <> oldSymbol Then
        
        'Print Ticker Symbol
        Cells(printTickerRow, 9).Value = oldSymbol
        
        'Print total vol
        Cells(printTickerRow, 10).Value = totalVol
                        
        'Location/Row for new symbol to be printed
        printTickerRow = printTickerRow + 1
     
        'Set total vol to first value of new Symbol
        totalVol = nextVol
              
    ' New ticker symbol the same
    Else
        
        ' Add/update total volume
        totalVol = totalVol + nextVol
                     
    End If
    
    Next i

End Sub



