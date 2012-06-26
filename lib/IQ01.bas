Attribute VB_Name = "IQ01"
Sub create()
    Sap.Go "iq01"
    Session.findById("wnd[0]/usr/ctxtRISA0-MATNR").text = assy_mrp
    Session.findById("wnd[0]/usr/ctxtRISA0-SERNR").text = serials
    Sap.Enter
    
    errorStr = "Serial number " & serials & " already exists for material " & assy_mrp
    errorBox = Session.findById("wnd[0]/sbar").text
    
    If errorStr = errorBox Then
        Call data.write_cell(6, i, "Already exists")
        Exit Sub
    Else
        Session.findById("wnd[0]/tbar[1]/btn[41]").press
        If (manuf_name & manuf_part) <> 0 Then
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\01").Select
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102C:SAPLITO0:1022/txtITOB-HERST").text = manuf_name
            Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102C:SAPLITO0:1022/txtITOB-MAPAR").text = manuf_part
        End If
        Sap.Save
        Call data.write_cell(6, i, "Material created")
    End If
End Sub
