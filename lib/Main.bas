Attribute VB_Name = "Main"
Public i As Integer

Sub Init()
    Sap.hook
    Excel_sheet.hook
    Color.Reset
End Sub

Sub multi(action As String)
    Main.Init
    
    For i = 4 To objSheet.UsedRange.Rows.Count
        Call Color.Yellow(i)
        Call Template.switch_case("read", action)
        Call Template.switch_case("execute", action)
        Call Color.Green(i)
        Excel_sheet.Scan_end
    Next
End Sub

