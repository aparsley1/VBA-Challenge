Attribute VB_Name = "Module1"
Sub StockData():
' Loop through all worksheets
     For Each ws In ThisWorkbook.Worksheets
' set variables
        Dim ticker As String
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalvolume As Double
        Dim LastRow As Long
        Dim summaryrow As Long
        Dim openingprice As Double
        Dim closingprice As Double
 'set column headers for summary
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
'find last row in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
' set summary table row
        summaryrow = 2
' loop through all data
    For i = 2 To LastRow
' see if ticker has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'get the variables
            ticker = ws.Cells(i, 1).Value
            
            openingprice = ws.Cells(i, 3).Value
            
            closingprice = ws.Cells(i, 6).Value
            
            yearlychange = closingprice - openingprice
'calculate percent change
            If openingprice <> 0 Then
            
                percentchange = (yearlychange / openingprice) * 100
            
            Else
                percentchange = 0
            End If
            
'add variables to summary
            ws.Cells(summaryrow, 9).Value = ticker
            ws.Cells(summaryrow, 10).Value = yearlychange
            ws.Cells(summaryrow, 11).Value = percentchange
            ws.Cells(summaryrow, 12).Value = totalvolume
'format percent change as a percentage
            ws.Cells(summaryrow, 11).NumberFormat = "0.00%"

'conditional formatting
            If yearlychange > 0 Then
                ws.Cells(summaryrow, 10).Interior.Color = vbGreen
        
            ElseIf yearlychange < 0 Then
                ws.Cells(summaryrow, 10).Interior.Color = vbRed
            End If
        
            summaryrow = summaryrow + 1
        
            totalvolume = 0
        
        End If
        
'add total volume for current ticker
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        
    Next i
    
' find greatest % increase
    Dim increaseticker As String
    Dim increasechange As Double
        
    For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        If ws.Cells(i, 11).Value > increasechange Then
        
        increasechange = ws.Cells(i, 11).Value
        
        increaseticker = ws.Cells(i, 9).Value
        
        End If
        
    Next i
    
    Next ws
'find greatest % decrease
    Dim decreseticker As String
    Dim decresechange As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
    For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
    
        If ws.Cells(i, 11).Value < decreasechange Then
        
        decreasechange = ws.Cells(i, 11).Value
        
        decreaseticker = ws.Cells(i, 9).Value
        
        End If
        
    Next i
    
    Next ws
    
    Dim greatestvolumeticker As String
    Dim greatestvolume As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
    For i = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    If ws.Cells(i, 12).Value > greatestvolume Then
    
    greatestvolume = ws.Cells(i, 12).Value
    
    End If
    
    Next i
    
    Next ws
    
    
End Sub



























