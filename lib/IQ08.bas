Attribute VB_Name = "IQ08"
Sub change()
    'Init regex to remove number from string
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = "[^A-Z]+"
    RegEx.IgnoreCase = True
    
    'Enter Equipment
    Sap.Go "iq08"
    Session.findById("wnd[0]/usr/txtSERNR-LOW").text = serials

    'Used to avoid TRACKING to appear when runnning IQ08
    Session.findById("wnd[0]/usr/ctxtSTAE1-LOW").text = "AVLB"
    
    Sap.Execute
    
    'Enter Sub Equipment
    Session.findById("wnd[0]/mbar/menu[4]/menu[8]").Select
    
    'Scan item description and build inputbox text
    popup = "Enter the number corresponding to the Serial: " & assy_serial & vbLf & vbLf
    colCount = Session.findById("wnd[0]/usr/subEQUIPMENTS:SAPLIEL2:0110/tblSAPLIEL2TCTRL_0110").Columns(6).Count
    For item_in = 0 To (colCount - 1)
        itemd = Session.findById("wnd[0]/usr/subEQUIPMENTS:SAPLIEL2:0110/tblSAPLIEL2TCTRL_0110/txtIEQINSTALL-EQKTX[6," & item_in & "]").text
        trimed = RegEx.Replace(itemd, "") 'Remove number from string
        popup = popup & item_in & " - " & trimed & vbLf & vbLf 'Concatenate every line in a single string
    Next
    
    'Ask wich item to use
    Index = Application.InputBox(popup, "Item Selection")
   
    'Select the requested line
    Session.findById("wnd[0]/usr/subEQUIPMENTS:SAPLIEL2:0110/tblSAPLIEL2TCTRL_0110").getAbsoluteRow(Index).Selected = True
    Sap.Execute
    Session.findById("wnd[1]").sendVKey 4
    
    'Input for replacement serial
    Session.findById("wnd[0]/usr/txtSERNR-LOW").text = assy_serial
    Sap.Execute
    
    'Save and go back to notification
    Session.findById("wnd[1]/tbar[0]/btn[16]").press
    Sap.Back
    Sap.Save
End Sub

Sub partout()
    Sap.Go "iq08"
    Session.findById("wnd[0]/usr/txtSERNR-LOW").text = serial
    Session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = part
    Session.findById("wnd[0]").sendVKey 8
    
    Session.findById("wnd[0]").sendVKey 39
    
    If Session.findById("wnd[0]/sbar").text <> "" Then Sap.Enter
    
    Session.findById("wnd[1]/usr/ctxtRISA0-MATNEU").text = IW72.partout
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    '2017 HDD, RAM
    If part = "T2017#A/400-01" Or part = "T2017#B/400-01" Or part = "T2017#C/400-01" Then
        Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08").Select
        If Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/chkRISA0-KRFKZ").Selected = True Then
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/btnCOPY").press
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/btnMAINTAIN").press
        Else
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/btnMAINTAIN").press
        End If
        Session.findById("wnd[0]/usr/subCHARACTERISTICS:SAPLCEI0:1400/tblSAPLCEI0CHARACTER_VALUES/ctxtRCTMS-MNAME[0,1]").SetFocus
        Session.findById("wnd[0]").sendVKey 2
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,3]").Selected = True
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,3]").SetFocus
        Session.findById("wnd[1]/tbar[0]/btn[8]").press
        Session.findById("wnd[0]/usr/subCHARACTERISTICS:SAPLCEI0:1400/tblSAPLCEI0CHARACTER_VALUES/ctxtRCTMS-MNAME[0,4]").SetFocus
        Session.findById("wnd[0]").sendVKey 2
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,2]").Selected = True
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,2]").SetFocus
        Session.findById("wnd[1]/tbar[0]/btn[8]").press
        Sap.Back
    End If
    '2018 HDD, CF 1GB
    If part = "T2018AAA/850-03" Or part = "T2018ABA/850-03" Or part = "T2018ACA/850-03" Then
        Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08").Select
        If Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/chkRISA0-KRFKZ").Selected = True Then
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/btnCOPY").press
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/btnMAINTAIN").press
        Else
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\08/ssubSUB_DATA:SAPMIEQ0:1500/btnMAINTAIN").press
        End If
        Session.findById("wnd[0]/usr/subCHARACTERISTICS:SAPLCEI0:1400/tblSAPLCEI0CHARACTER_VALUES/ctxtRCTMS-MNAME[0,0]").SetFocus
        Session.findById("wnd[0]").sendVKey 2
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,3]").Selected = True
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,3]").SetFocus
        Session.findById("wnd[1]/tbar[0]/btn[8]").press
        Session.findById("wnd[0]/usr/subCHARACTERISTICS:SAPLCEI0:1400/tblSAPLCEI0CHARACTER_VALUES/ctxtRCTMS-MNAME[0,1]").SetFocus
        Session.findById("wnd[0]").sendVKey 2
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,3]").Selected = True
        Session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,3]").SetFocus
        Session.findById("wnd[1]/tbar[0]/btn[8]").press
        Sap.Back
    End If
    Sap.Save
End Sub
