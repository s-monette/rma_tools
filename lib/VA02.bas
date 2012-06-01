Attribute VB_Name = "VA02"
Sub swap()
    Sap.Go "VA02"
    Title = Session.findById("wnd[0]/titl").text 'Copy Title of transaction for error handling loop
    Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = rma
        
    Sap.Enter
    While Session.findById("wnd[0]/titl").text = Title 'Error handling of transaction still under process
        Sap.Enter
    Wend
    
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_POPO").press
    Session.findById("wnd[1]/usr/txtRV45A-POSNR").text = Item
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ZZTIMER_TWO[0,0]").SetFocus
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_PREP").press
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/btnBT_RALL").press
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/btnBT_RINS").press
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/tblSAPLV46RTCTRL_REPPO/chkV46R_ITEM-VORGA_VAL_106[4,2]").Selected = True
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/tblSAPLV46RTCTRL_REPPO/txtV46R_ITEM-MENGE[0,2]").text = "1"
    Sap.Back
    Session.findById("wnd[0]/tbar[1]/btn[17]").press
    Session.findById("wnd[1]/usr/sub:SAPLATP4:0600/chkRV03V-SELKZ[0,0]").Selected = True
    Session.findById("wnd[1]/tbar[0]/btn[5]").press
    Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
    Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
    Session.findById("wnd[0]/tbar[1]/btn[35]").press
    Session.findById("wnd[0]/usr/btnBUT2").press
    
    If Worksheets("Shipping").chkStock.Value = True Then 'Set the swap in location 0015
        Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_MKAL").press
        Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_MKLO").press
        Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_POPO").press
        Session.findById("wnd[1]/usr/txtRV45A-POSNR").text = Item
        Session.findById("wnd[1]").sendVKey 0
        Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[4,3]").SetFocus
        Session.findById("wnd[0]/mbar/menu[2]/menu[2]/menu[2]").Select
        Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03").Select
        Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03/ssubSUBSCREEN_BODY:SAPMV45A:4452/ctxtVBAP-LGORT").text = "0015"
        Sap.Back
    End If
    Sap.Save
End Sub

Sub partout()
    Sap.Go "va02"

    Title = Session.findById("wnd[0]/titl").text 'Copy Title of transaction for error handling loop
    Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = rma
    Sap.Enter
    
    While Session.findById("wnd[0]/titl").text = Title 'Error handling of transaction still under process
        Sap.Enter
    Wend
    
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_POPO").press
    Session.findById("wnd[1]/usr/txtRV45A-POSNR").text = Item
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ZZTIMER_TWO[0,0]").SetFocus
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_PREP").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/tblSAPLV46RTCTRL_REPPO/chkV46R_ITEM-VORGA_VAL_103[1,0]").Selected = True
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/tblSAPLV46RTCTRL_REPPO").getAbsoluteRow(0).Selected = True
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/tblSAPLV46RTCTRL_REPPO/txtV46R_ITEM-MENGE[0,0]").text = "1"
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/tblSAPLV46RTCTRL_REPPO/ctxtV46R_ITEM-MATNR_G[4,0]").text = IW72.partout
    Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV46R:4100/btnBT_RSER").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/btnBT_SSEA").press
    Session.findById("wnd[1]/tbar[0]/btn[5]").press
    Sap.Back
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Sap.Save
End Sub
