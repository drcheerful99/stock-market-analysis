Sub stockAnalysis()
    ' declare variables
    Dim ws As Worksheet
    Dim Ticker As String
    Dim annualChange As Double
    Dim pctChange As Double
    Dim totalVolume As Double
    ' Dim startValue As Double
    Dim lastRow As Integer
    
    ' loop through all worksheets in current active workbook
    For Each ws In Worksheets
    
    ' set variable values for each worksheet
    j = 0
    totalVolume = 0
    annualChange = 0
    StartValue = 2
    
    ' add headers
    ws.Range("i1").Value = "Ticker Symbol"
    ws.Range("j1").Value = "Annual Change"
    ws.Range("k9").Value = "Percent Change"
    ws.Range("l1").Value = "Total Volume"
    
    ' find row # of the last row with data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            ' check whether stock symbol is the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' sum up stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            If totalVolume = 0 Then
                ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = 0 & "%"
                ws.Range("l" & 2 + j).Value = 0
            Else
                If ws.Cells(StartValue, 3) = 0 Then
                    For findValue = StartValue To i
                        If ws.Cells(findValue, 3).Value <> 0 Then
                            StartValue = findValue
                    Exit For
                        End If
                    Next findValue
                End If
                
            ' calculations
            annualChange = (ws.Cells(i, 6) - ws.Cells(StartValue, 3))
            pctChange = Round((annualChange / ws.Cells(StartValue, 3) * 100), 2)
            
            StartValue = i + 1
            ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("j" & 2 + j).Value = Round(annualChange, 2)
            ws.Range("k" & 2 + j).Value = pctChange & "%"
            ws.Range("l" & 2 + j).Value = totalVolume
            
            ' set cell color based on positive/negative pct change
            If annualChange > 0 Then
              ws.Range("j" & 2 + j).Interior.ColorIndex = 4
            ElseIf annualChange < 0 Then ws.Range("j" & 2 + j).Interior.ColorIndex = 3
          End If
          
           
            annualChange = 0
            totalVolume = 0
            j = j + 1
            
        Else
                totalVolume = totoalVolume + ws.Cells(i, 7).Value
                
            End If
        Next i
    Next ws
End Sub