Attribute VB_Name = "Excel_sheet"
Public objSheet As Object

Sub hook()
    Set objSheet = ThisWorkbook.Worksheets("Shipping")
End Sub

Sub Scan_end()
    'Check if next line is empty
    countRow = Application.CountA(Sheets("Shipping").Rows(i + 1))
    If countRow = 0 Then
        Sap.Go ""
        Call data.write_cell(1, i + 1, "Finished")
        End
    End If
End Sub
