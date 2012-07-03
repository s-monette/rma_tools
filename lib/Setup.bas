Attribute VB_Name = "Setup"
Sub Export_bas()
    For Each objFile In ThisWorkbook.VBProject.VBComponents
        If objFile.Type = "1" Then objFile.export ThisWorkbook.Path & "\lib\" & objFile.Name & ".bas"
    Next
End Sub

Sub Init()
    Template.dropBox
    Excel_sheet.hook
    Color.Reset
End Sub
