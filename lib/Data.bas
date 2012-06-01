Attribute VB_Name = "Data"
Sub Read()
    'Get data from Excel to variables
    Call Template.switch_case("read", Worksheets("Shipping").cmbInput.text)
End Sub

Sub write_cell(ByVal x, ByVal y As Integer, inputStr As String)
    objSheet.Cells(y, x).Value = inputStr
End Sub

Sub clipboard(text As String)
    Set data_string = New DataObject
    data_string.SetText text
    data_string.PutInClipboard
End Sub

Sub Serial_batch()
    itt = 4
    While objSheet.Cells(itt, 1).Value <> ""
        itt = itt + 1
    Wend
    Range("A4:A" & (itt - 1)).Copy
End Sub
