Attribute VB_Name = "Setup"
Sub Export_bas()
    For Each objFile In ThisWorkbook.VBProject.VBComponents
        If objFile.Type = "1" Then objFile.export ThisWorkbook.Path & "\lib\" & objFile.Name & ".bas"
    Next
End Sub

Sub RemoveAllMacros()
    For Each objFile In ThisWorkbook.VBProject.VBComponents
        If objFile.Type = "1" Then ThisWorkbook.VBProject.VBComponents.Remove objFile
    Next
End Sub

Sub Import_bas()
    Set FileSys = CreateObject("Scripting.fileSystemObject")
    Set libFolder = FileSys.GetFolder(ThisWorkbook.Path & "\lib\")
  
    For Each objFile In libFolder.Files
        ThisWorkbook.VBProject.VBComponents.Import objFile
    Next
End Sub

Sub Init()
    Template.dropBox
    Excel_sheet.hook
    Color.Reset
End Sub
