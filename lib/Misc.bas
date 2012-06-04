Attribute VB_Name = "Misc"
Sub Spool()
    Main.Init
    Sap.Go ("SP02")
    Session.findById("wnd[0]/tbar[1]/btn[48]").press 'Select all
    Session.findById("wnd[0]/tbar[1]/btn[44]").press 'Print selected
End Sub

Sub home()
    Main.Init
    serials = ""
    Sap.Go ("iw72")
    IW72.Config
    End
End Sub

Sub Zkciresrep()
    Sap.Go ("ZKCIRESREP")
    Session.findById("wnd[0]/usr/txtS_TEST").text = req_num
    Session.findById("wnd[0]/usr/ctxtP_BUKRS").text = "1000"
    Session.findById("wnd[0]/usr/ctxtP_WERKS").text = "1000"
    Sap.Execute
    
    Session.findById("wnd[0]/tbar[0]/btn[86]").press
    Session.findById("wnd[1]/tbar[0]/btn[13]").press
End Sub

Sub check_row()
    Excel_sheet.hook
    If ActiveWindow.ScrollRow < 5 Then
        Call data.write_cell(1, 2, "At first row")
        Call Color.Green(2, 1)
    Else
        Call data.write_cell(1, 2, "Not at first row")
        Call Color.Yellow(2, 1)
    End If
End Sub

Sub check_version()
    If Dir(ThisWorkbook.Path & "\lib\") <> "" Then Setup.Export_bas
    If Dir(ThisWorkbook.Path & "\.git\") <> "" Then
        sFile = ThisWorkbook.Path & "\.git\refs\heads\master"
        Open sFile For Input As FreeFile
            sString = Input$(LOF(1), 1)
        Close
        gitHash = Left(sString, Len(sString) - 1)
        currentHash = ThisWorkbook.BuiltinDocumentProperties("Comments").Value
        currentRev = ThisWorkbook.BuiltinDocumentProperties("Revision number").Value
        If currentHash <> gitHash Then
            ThisWorkbook.BuiltinDocumentProperties("Revision number") = currentRev + 1
            ThisWorkbook.BuiltinDocumentProperties("Comments") = gitHash
        End If
    End If
End Sub
