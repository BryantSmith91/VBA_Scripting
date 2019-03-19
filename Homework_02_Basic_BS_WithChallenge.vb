Sub HomeworkBasic()
Dim ws As Worksheet
    
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    Dim ticker As String
    Dim totalvolume As Double
    Dim output As Integer
    Dim lrow As Double

        output = 2
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Total Volume"
        lrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lrow
            ticker = Cells(i, 1)
            If ticker <> Cells(i + 1, 1) Then
                totalvolume = totalvolume + Cells(i, 7)
                Cells(output, 9) = ticker
                Cells(output, 10).Value = totalvolume
                output = output + 1
                totalvolume = 0
                
            ElseIf ticker = Cells(i + 1, 1) Then
                totalvolume = totalvolume + Cells(i, 7)
            
            ElseIf Cells(i, 1) = " " Then
                Exit For
            
            End If
            
        Next i
	Next ws

End Sub


