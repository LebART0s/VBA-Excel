Sub CheckNum()
	Dim intRow
	For intRow = 1 To 20
		If Cells(intRow, 2).Value = Cells(1, 2).Value Then
			Range(Cells(intRow, 2), Cells(intRow, 2)).Interior.Color = RGB(255, 0, 0)	
		Else
		Range(Cells(intRow, 2), Cells(intRow, 2)).Interior.Pattern = xlNone
		Range(Cells(intRow, 2), Cells(intRow, 2)).Interior.TintAndShade = 0
		Range(Cells(intRow, 2), Cells(intRow, 2)).Interior.PatternTintAndShade = 0
		If Cells(intRow, 3).Value = Cells(1, 2).Value Then
			Range(Cells(intRow, 3), Cells(intRow, 3)).Interior.Color = RGB(255, 0, 0)
		Else
			Range(Cells(intRow, 3), Cells(intRow, 3)).Interior.Pattern = xlNone
			Range(Cells(intRow, 3), Cells(intRow, 3)).Interior.TintAndShade = 0
			Range(Cells(intRow, 3), Cells(intRow, 3)).Interior.PatternTintAndShade = 0
			If Cells(intRow, 4).Value = Cells(1, 2).Value Then
				Range(Cells(intRow, 4), Cells(intRow, 4)).Interior.Color = RGB(255, 0, 0)
			Else
				Range(Cells(intRow, 4), Cells(intRow, 4)).Interior.Pattern = xlNone
				Range(Cells(intRow, 4), Cells(intRow, 4)).Interior.TintAndShade = 0
				Range(Cells(intRow, 4), Cells(intRow, 4)).Interior.PatternTintAndShade = 0
				If Cells(intRow, 5).Value = Cells(1, 2).Value Then
					Range(Cells(intRow, 5), Cells(intRow, 5)).Interior.Color = RGB(255, 0, 0)
				Else
					Range(Cells(intRow, 5), Cells(intRow, 5)).Interior.Pattern = xlNone
					Range(Cells(intRow, 5), Cells(intRow, 5)).Interior.TintAndShade = 0
					Range(Cells(intRow, 5), Cells(intRow, 5)).Interior.PatternTintAndShade = 0
					If Cells(intRow, 6).Value = Cells(1, 2).Value Then
						Range(Cells(intRow, 6), Cells(intRow, 6)).Interior.Color = RGB(255, 0, 0)
					Else
						Range(Cells(intRow, 6), Cells(intRow, 6)).Interior.Pattern = xlNone
						Range(Cells(intRow, 6), Cells(intRow, 6)).Interior.TintAndShade = 0
						Range(Cells(intRow, 6), Cells(intRow, 6)).Interior.PatternTintAndShade = 0
						If Cells(intRow, 7).Value = Cells(1, 2).Value Then
							Range(Cells(intRow, 7), Cells(intRow, 7)).Interior.Color = RGB(255, 0, 0)
						Else
							CellFill(intRow,7)
						End If
					End If
				End If
			End If
		End If
	End If
	Next intRow
End Sub

Sub CellFill(sintRow, sintCol)
	Range(Cells(sintRow, sintCol), Cells(sintRow, sintCol)).Interior.Pattern = xlNone
	Range(Cells(sintRow, sintCol), Cells(sintRow, sintCol)).Interior.TintAndShade = 0
	Range(Cells(sintRow, sintCol), Cells(sintRow, sintCol)).Interior.PatternTintAndShade = 0
End Sub