Sub HomeworkMed()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

	Dim ticker As String
	Dim totalvolume, OpenYear, CloseYear, YearlyChange, PercentChange, lrow As Long
	Dim output As Integer


		output = 2
		Cells(1, 9) = "Ticker"
		Cells(1, 10) = "Yearly Change"
		Cells(1, 11) = "Percent Change"
		Cells(1, 12) = "Total Volume"

		lrow = Cells(Rows.Count, 1).End(xlUp).Row

		OpenYear = Cells(2, 3).Value

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
				totalvolume = 0
				OpenYear = Cells(i + 1, 3).Value

					If YearlyChange >= 0 Then
					Cells(output, 10).Interior.ColorIndex = 4

					ElseIf YearlyChange < 0 Then
					Cells(output, 10).Interior.ColorIndex = 3

					End If



				output = output + 1

			ElseIf ticker = Cells(i + 1, 1) Then
				totalvolume = totalvolume + Cells(i, 7)

			ElseIf Cells(i, 1) = " " Then
				Exit For

			End If

		Next i
	Next ws

End Sub
