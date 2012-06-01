Attribute VB_Name = "Color"
Sub Green(line As Integer, Optional col As Integer = 1)
    objSheet.Cells(line, col).Interior.ColorIndex = 4 'Set the cell green
End Sub

Sub Yellow(line As Integer, Optional col As Integer = 1)
    objSheet.Cells(line, col).Interior.ColorIndex = 6 'Set the cell yellow
End Sub

Sub Reset()
    objSheet.Cells.Interior.ColorIndex = 0 'Reset the color of all cells to blank
    Misc.check_row
End Sub
