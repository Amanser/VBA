Sub Button1_Click()





' Declare sht as a worksheet.
Dim sht As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each sht In ThisWorkbook.Worksheets

    
Dim i, x As Variant
Dim lastRow As Variant
Dim SummaryRow As Variant
Dim volume As Variant
Dim ticker As String
Dim openprice As Double
Dim closeprice As Double



'Find number of rows in the sheet
lastRow = sht.Cells(Rows.Count, 1).End(xlUp).Row


SummaryRow = 1
volume = 0

'Add column headers
sht.Cells(1, 9).Value = "Ticker"
sht.Cells(1, 10).Value = "Total Stock Volume"
'sht.Cells(1, 11).Value = "Yearly Change"
'sht.Cells(1, 12).Value = "Percentage Change"

    For i = 2 To lastRow

    volume = volume + sht.Cells(i, 7).Value

        'If sht.Cells(i, 2).Value = 20160101 Or sht.Cells(i, 2).Value = 20140101 Then
        '    openprice = sht.Cells(i, 3).Value
            
       ' ElseIf sht.Cells(i, 2).Value = 20161230 Or sht.Cells(i, 2).Value = 20141231 Then
       '     closeprice = sht.Cells(i, 6).Value
       ' End If
        
        
        If sht.Cells(i, 1).Value <> sht.Cells(i + 1, 1).Value Then
        
            ticker = sht.Cells(i, 1).Value
        
            SummaryRow = SummaryRow + 1
                    
            sht.Cells(SummaryRow, 9).Value = ticker
        
            sht.Cells(SummaryRow, 10).Value = volume
            
           ' sht.Cells(SummaryRow, 11).Value = (closeprice - openprice)
            
               ' If closeprice <> 0 Then
                        
                'sht.Cells(SummaryRow, 12).Value = (closeprice - openprice) / closeprice
        
                'Else
            
               ' sht.Cells(SummaryRow, 12).Value = "Divide by 0 error"
            
               ' End If
            
            
           ' openprice = 0
                       
            'closeprice = 0
        
            volume = 0
        
        End If

    Next i


    Next sht








End Sub
