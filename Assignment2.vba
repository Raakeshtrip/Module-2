Sub getWorkSheetName()
    Dim ws As Worksheet
    Dim i As Integer
    
' Get each sheet name and pass on to calculate function
    
    Dim worksheetName As String
    i = 1
    
    For Each ws In Worksheets
        worksheetName = ws.Name
        calculateTotal (ws.Name)
        Next ws
        
    End Sub
    
    
    Sub calculateTotal(ByVal sheetName As String)
        
'Declare all varibale needed for calculation
        Dim shtname As Worksheet ' sheet name
        
        Dim tickername As String 'ticker name
        
        Dim totalvol As Double   'total volume for each ticker
        
        Dim lengthOfSheet As Long ' max rows of each sheet
        
        Dim lengthResult As Integer ' length of the cells where ticker and volume stored
        
        
'Counter variables
        
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim min As Double
        Dim max As Double
        
        
        
' Variables to store the results
        
        Dim maxValue As Double
        Dim minValue As Double
        Dim totalVolume As Double
        Dim greatestVol As Double
        Dim greatestVolTckName As String
        Dim maxPercentageOfIncrease As Double
        Dim maxPercentageOfIncreasetckername As String
        Dim maxPercentageOfDecrease As Double
        Dim mAxPercentageOfDecreasetckername As String
        
'Intitalize all the variables
        
        greatestVolTckName = 0
        maxPercentageOfIncrease = 0
        maxPercentageOfIncreasetckername = 0
        maxPercentageOfDecrease = 0
        mAxPercentageOfDecreasetckername = 0
        
        totalVolume = 0
        
        i = 1
        
        j = 2
        
        k = 2
        
        lengthOfSheet = 0
        
        lengthResult = 0
        
        totalvol = 0
        
'Set the current worksheet name and max row length
        Set shtname = ActiveWorkbook.Worksheets(sheetName)
        lengthOfSheet = shtname.Range("A" & Rows.Count).End(xlUp).Row
'For each loop accross each row in the sheet
        For i = 2 To lengthOfSheet
            tickername = shtname.Cells(i, 1).Value 'Get the current ticker name
            totalvol = totalvol + shtname.Cells(i, 7).Value ' Calculate volume by adding each ticker row by row
            If (shtname.Cells(i + 1, 1) <> tickername) Then ' Checkthe current ticker variable as compared next one if its different then  update all the values
                min = shtname.Cells(k, 3).Value
                   k = i + 1
                shtname.Cells(j, 9).Value = tickername
                shtname.Cells(j, 10).Value = totalvol
                max = shtname.Cells(k - 1, 6).Value
                shtname.Cells(j, 13).Value = max - min
               
             shtname.Cells(j, 13).NumberFormat = "0.0000000000"
                
                If (shtname.Cells(j, 13).Value >= 0) Then
                    shtname.Cells(j, 13).Interior.ColorIndex = 4
                Else
                    shtname.Cells(j, 13).Interior.ColorIndex = 3
                End If
                If (min <> 0) Then
                    totalVolume = shtname.Cells(j, 13).Value / min
                    
                Else
                    shtname.Cells(j, 14).Value = 0
                    shtname.Cells(j, 14).NumberFormat = "0.00%"
                   
                End If
                shtname.Cells(j, 14).Value = totalVolume
                shtname.Cells(j, 14).NumberFormat = "0.00%"
               
                j = j + 1
'set the maximum total value for one tcker
                If (totalvol > greatestVol) Then
                    greatestVol = totalvol
                    greatestVolTckName = tickername
                End If
                If (totalVolume > maxPercentageOfIncrease) Then
                    maxPercentageOfIncrease = totalVolume
                    maxPercentageOfIncreasetckername = tickername
                End If
                If (totalVolume < maxPercentageOfDecrease) Then
                    maxPercentageOfDecrease = totalVolume
                    mAxPercentageOfDecreasetckername = tickaername
                End If
                totalvol = 0
            End If
            Next i
            
            shtname.Cells(1, 9).Value = "Ticker Name"
            shtname.Cells(1, 10).Value = "Volume"
            shtname.Cells(1, 13) = "Yearly Change"
            shtname.Cells(1, 14) = "Percent Change"
            
            shtname.Cells(4, 16) = "Greatest Total volume"
            
            shtname.Cells(4, 18) = greatestVol
            
            shtname.Cells(4, 17) = greatestVolTckName
            
            
            shtname.Cells(3, 16) = "Greater % of Derease"
            
            shtname.Cells(3, 18) = maxPercentageOfDecrease
            
            shtname.Cells(3, 18).NumberFormat = "0.00%"
             
            
            shtname.Cells(3, 17) = mAxPercentageOfDecreasetckername
            
            
            
            
            shtname.Cells(2, 16) = "Greater % of Increase"
            
            shtname.Cells(2, 18) = maxPercentageOfIncrease
            shtname.Cells(2, 18).NumberFormat = "0.00%"
            
            shtname.Cells(2, 17) = maxPercentageOfIncreasetckername
            
            
            shtname.Cells(1, 17).Value = "Ticker"
            shtname.Cells(1, 18).Value = "Value"
            
            
        End Sub
        
        
        
        
        
        
        
        


