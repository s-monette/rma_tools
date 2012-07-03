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

'The code below is contain in ThisWorkBook and cannot be tracked with GIT.
'Added here for reference.

'Sub RemoveAllMacros()
'    For Each objFile In ThisWorkbook.VBProject.VBComponents
'        If objFile.Type = "1" Then ThisWorkbook.VBProject.VBComponents.Remove objFile
'    Next
'End Sub

'Sub Import_bas()
'    Set FileSys = CreateObject("Scripting.fileSystemObject")
'    Set libFolder = FileSys.GetFolder("\\D101000577\rma_tools\lib\")
'
'    For Each objFile In libFolder.Files
'        ThisWorkbook.VBProject.VBComponents.Import objFile
'    Next
'End Sub

'Sub net_update()
'    Application.ScreenUpdating = True
'    If Dir("\\D101000577\rma_tools\.git\") <> "" Then
'        sFile = "\\D101000577\rma_tools\.git\refs\heads\master"
'        Open sFile For Input As FreeFile
'            sString = Input$(LOF(1), 1)
'        Close
'        gitHash = Left(sString, Len(sString) - 1)
'
'        sFile = "\\D101000577\rma_tools\version.rev"
'        Open sFile For Input As FreeFile
'            sString = Input$(LOF(1), 1)
'        Close
'        currentRev = Left(sString, Len(sString) - 2)
'        currentHash = ThisWorkbook.BuiltinDocumentProperties("Comments").Value
'        If gitHash <> currentHash Then
'            ThisWorkbook.BuiltinDocumentProperties("Comments") = gitHash
'            ThisWorkbook.BuiltinDocumentProperties("Revision number") = currentRev
'            ThisWorkbook.RemoveAllMacros
'            ThisWorkbook.Import_bas
'            MsgBox "Updated latest revision."
'        Else: MsgBox "No updated available."
'        End If
'    End If
'End Sub
