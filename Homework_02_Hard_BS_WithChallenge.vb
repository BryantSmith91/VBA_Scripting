Sub HomeworkHard()

Dim ws As Worksheet
    
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    Dim ticker, grinct, grdect, grtvt As String
    Dim totalvolume, OpenYear, CloseYear, YearlyChange, PercentChange, GrInc, GrDec, GrTv, lrow As Double
    Dim output, output2 As Integer
    
    
        output = 2
        output2 = 2
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Volume"
        Cells(1, 17) = "Ticker"
        Cells(1, 18) = "Value"
        Cells(2, 16) = "Greatest % Increase"
        Cells(3, 16) = "Greatest % Decrease"
        Cells(4, 16) = "Greatest Total Volume"
        
        GrInc = 0
        GrDec = 0
        GrTv = 0
        
        OpenYear = Cells(2, 3).Value
		
		lrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lrow
            ticker = Cells(i, 1)
            If ticker <> Cells(i + 1, 1) Then
                totalvolume = totalvolume + Cells(i, 7)
                Cells(output, 9) = ticker
                Cells(output, 12).Value = totalvolume
                CloseYear = Cells(i, 6).Value
                YearlyChange = CloseYear - OpenYear
                If OpenYear = 0 Then PercentChange = YearlyChange Else PercentChange = YearlyChange / OpenYear
                Cells(output, 10).Value = YearlyChange
                Cells(output, 11).Value = PercentChange
                Cells(output, 11).NumberFormat = "0.00%"
    
                OpenYear = Cells(i + 1, 3).Value
                    
                    If YearlyChange >= 0 Then
                        Cells(output, 10).Interior.ColorIndex = 4
                    
                        ElseIf YearlyChange < 0 Then
                        Cells(output, 10).Interior.ColorIndex = 3
                    
                    End If
                    
                If GrInc < PercentChange Then
                    GrInc = PercentChange
                    grinct = ticker
                    
                    ElseIf GrDec > PercentChange Then
                        GrDec = PercentChange
                        grdect = ticker
                End If
                
                If GrTv < totalvolume Then
                    GrTv = totalvolume
                    grtvt = ticker
                End If
                            
                
                totalvolume = 0
                output = output + 1
                    
            ElseIf ticker = Cells(i + 1, 1) Then
                totalvolume = totalvolume + Cells(i, 7)
            
            ElseIf Cells(i, 1) = " " Then
                Exit For
            
            End If
           
        Next i
         
        Cells(2, 17) = grinct
        Cells(2, 18).NumberFormat = "0.00%"
        Cells(3, 17) = grdect
        Cells(3, 18).NumberFormat = "0.00%"
        Cells(4, 17) = grtvt
        Cells(2, 18).Value = GrInc
        Cells(3, 18).Value = GrDec
        Cells(4, 18).Value = GrTv
        ActiveSheet.Columns("A:R").AutoFit
	Next ws

End Sub







