Attribute VB_Name = "Module1"
Sub YearStock()

Dim workSheetName As Worksheet
Dim endRow As Long
Dim endRow2 As Long
Dim tickerCount As Long
Dim tickerCount2 As Long
Dim tickerCount3 As Long
Dim infoDate As String
Dim infoDate2 As String

For Each workSheetName In ThisWorkbook.Worksheets
    
    
    endRow = workSheetName.Cells(Rows.Count, 1).End(xlUp).Row
    'fill colum titles in the cells
    workSheetName.Cells(1, 9).Value = "Ticker"
    workSheetName.Cells(1, 10).Value = "Yearly change"
    workSheetName.Cells(1, 11).Value = "Percentage Change"
    workSheetName.Cells(1, 12).Value = "Total Stock Volume"
    workSheetName.Cells(2, 15).Value = "Greatest % Increase"
    workSheetName.Cells(3, 15).Value = "Greatest % Decrease"
    workSheetName.Cells(4, 15).Value = "Greatest Total Volume"
    workSheetName.Cells(1, 16).Value = "Ticker"
    workSheetName.Cells(1, 17).Value = "Value"
    'Following for clean cells
    For cleanCell = 2 To endRow
        workSheetName.Range("I" & cleanCell).Value = ""
        workSheetName.Range("J" & cleanCell).Value = ""
        workSheetName.Range("K" & cleanCell).Value = ""
        workSheetName.Range("L" & cleanCell).Value = ""
        workSheetName.Range("P" & cleanCell).Value = ""
        workSheetName.Range("Q" & cleanCell).Value = ""
    Next cleanCell
    'comparason per row
    workSheetName.Cells(2, 9).Value = workSheetName.Cells(2, 1).Value
    tickerCount2 = 2
    For tickerCount = 2 To endRow
            
        If workSheetName.Cells(tickerCount2, 9).Value <> workSheetName.Cells(tickerCount, 1).Value Then
            'fill colum I (Ticker)
            tickerCount2 = tickerCount2 + 1
            workSheetName.Cells(tickerCount2, 9).Value = workSheetName.Cells(tickerCount, 1).Value
                
        Else
            ' count rows in colum I
            For tickerCount3 = 2 To tickerCount2
                'fill colum L (Total Stock Volume)
                If workSheetName.Cells(tickerCount, 1).Value = workSheetName.Cells(tickerCount3, 9).Value Then
                
                    workSheetName.Cells(tickerCount3, 12).Value = workSheetName.Cells(tickerCount3, 12).Value + workSheetName.Cells(tickerCount, 7).Value
                
                End If
            Next tickerCount3
        End If
        
    Next tickerCount
    '**********************************************************
    'yearly change and percentage change
    For tickerCount3 = 2 To tickerCount2
        
        For tickerCount = 2 To endRow
            
            If workSheetName.Cells(tickerCount, 1).Value = workSheetName.Cells(tickerCount3, 9).Value Then
                
                infoDate = workSheetName.Name + "0102"
                infoDate2 = workSheetName.Name + "1231"
                'get the opening price
                If workSheetName.Cells(tickerCount, 2).Value = infoDate Then
                
                    openingPrice = workSheetName.Cells(tickerCount, 3).Value
                
                End If
                'get the closing price
                If workSheetName.Cells(tickerCount, 2).Value = infoDate2 Then
                
                    closingPrice = workSheetName.Cells(tickerCount, 6).Value
                
                End If
                
            End If
            
        Next tickerCount
        'yearly change
        workSheetName.Cells(tickerCount3, 10).Value = closingPrice - openingPrice
        'percentage change
        workSheetName.Cells(tickerCount3, 11).Value = workSheetName.Cells(tickerCount3, 10).Value / openingPrice
        'Format to the cells
        workSheetName.Cells(tickerCount3, 11).NumberFormat = "0.00%"
        If workSheetName.Cells(tickerCount3, 10).Value > 0 Then
        
            workSheetName.Cells(tickerCount3, 10).Interior.ColorIndex = 4
        Else
        
            workSheetName.Cells(tickerCount3, 10).Interior.ColorIndex = 3
        
        End If
        
    Next tickerCount3
        
    
    'Greatest Increase, Decrease and Total Volume
    greatestIncrease = workSheetName.Cells(2, 11).Value
    tickerIncrease = workSheetName.Cells(2, 9).Value
    greatestDecrease = workSheetName.Cells(2, 11).Value
    tickerDecrease = workSheetName.Cells(2, 9).Value
    greatestTotalVolume = workSheetName.Cells(2, 12).Value
    tickerTotalVolume = workSheetName.Cells(2, 9).Value
    For tickerCount3 = 2 To tickerCount2
    
        If workSheetName.Cells(tickerCount3, 11).Value > greatestIncrease Then
             
           greatestIncrease = workSheetName.Cells(tickerCount3, 11).Value
           tickerIncrease = workSheetName.Cells(tickerCount3, 9).Value
        End If
        
        If workSheetName.Cells(tickerCount3, 11).Value < greatestIncrease Then
             
           greatestDecrease = workSheetName.Cells(tickerCount3, 11).Value
           tickerDecrease = workSheetName.Cells(tickerCount3, 9).Value
            
        End If
        
        If workSheetName.Cells(tickerCount3, 12).Value > greatestTotalVolume Then
             
           greatestTotalVolume = workSheetName.Cells(tickerCount3, 12).Value
           tickerTotalVolume = workSheetName.Cells(tickerCount3, 9).Value
        End If
    
    Next tickerCount3
    
    workSheetName.Cells(2, 16).Value = tickerIncrease
    workSheetName.Cells(2, 17).Value = greatestIncrease
    
    workSheetName.Cells(3, 16).Value = tickerDecrease
    workSheetName.Cells(3, 17).Value = greatestDecrease
    
    workSheetName.Cells(4, 16).Value = tickerTotalVolume
    workSheetName.Cells(4, 17).Value = greatestTotalVolume
    
    
Next workSheetName

MsgBox ("Your data is ready ;)")

End Sub



