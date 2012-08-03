Attribute VB_Name = "IW72"
Public req_num, in_out, Order, rma, part, Item, serial, partout, batchout, batch As String

Sub Enter()
    'Enter the transaction
    Sap.Go "iw72"
    
    IW72.Config
    Sap.Execute
    
    On Error Resume Next
        Session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    
    'Check if function is currently locked
    While (Session.findById("wnd[0]/titl").text = "Change Order: Initial Screen")
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Sap.Execute
    Wend
End Sub

Sub Config()
    'Check checkbox incl.object list
    Session.findById("wnd[0]/usr/chkDY_OBL").Selected = True
    
    'Remove time period
    Session.findById("wnd[0]/usr/ctxtDATUV").text = "" 'Start
    Session.findById("wnd[0]/usr/ctxtDATUB").text = "" 'End
    
    'Enter Serial number and filter
    Session.findById("wnd[0]/usr/txtSERIALNR-LOW").text = serials
    Session.findById("wnd[0]/usr/ctxtGEWRK-LOW").text = "rma"
    Session.findById("wnd[0]/usr/txtVAWRK-LOW").text = "1010"
End Sub

Sub Info()
    'Get part# rma# item# and service order
    serial = serials 'Save values so they wont be lost when doing catalogue loop
    partout = partouts
    batchout = batchouts
    'Get board informations
    Order = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR").text
    rma = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subSUB_SERVICE:SAPLCOI3:0700/subSUB01:SAPLCOI3:0601/txtCAUFVD-RMANR").text
    part = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subSUB_SERVICE:SAPLCOI3:0700/subSUB01:SAPLCOI3:0601/ctxtCAUFVD-MATNR").text
    Item = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subSUB_SERVICE:SAPLCOI3:0700/subSUB01:SAPLCOI3:0601/txtCAUFVD-POSNV_RMA").text
End Sub

Sub Goto_object()
    'Go to Objet/notification
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIOLU").Select
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIOLU/ssubSUB_AUFTRAG:SAPLIWOL:0300/tblSAPLIWOLOBJK_120/btnRIWOL0-IMELD[10,0]").SetFocus
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIOLU/ssubSUB_AUFTRAG:SAPLIWOL:0300/tblSAPLIWOLOBJK_120/btnRIWOL0-IMELD[10,0]").press
End Sub

Sub get_batch()
    'Go to Objet/notification
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIOLU").Select
    'Get batch# from EQUI
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIOLU/ssubSUB_AUFTRAG:SAPLIWOL:0300/tblSAPLIWOLOBJK_120/ctxtRIWOL-SERNR[2,0]").SetFocus
    Session.findById("wnd[0]").sendVKey 2
    batch = Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122C:SAPLITO0:1220/ctxtITOB-CHARGE").text
    Sap.Back
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIOLU/ssubSUB_AUFTRAG:SAPLIWOL:0300/tblSAPLIWOLOBJK_120/btnRIWOL0-IMELD[10,0]").SetFocus
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIOLU/ssubSUB_AUFTRAG:SAPLIWOL:0300/tblSAPLIWOLOBJK_120/btnRIWOL0-IMELD[10,0]").press
End Sub

Sub repair_log()
    'Enter repair log, paste log text and go back
    Session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/btnQMICON-LTMELD").press
    data.clipboard (logs)
    Session.findById("wnd[0]/mbar/menu[1]/menu[2]").Select
    Sap.Back
    
    'Enter replaced component and catalogue loop
    flag_next = 1 'Used to determine if a new iteration is needed
    While (flag_next = 1)
        Color.Yellow (i)
        'Get data from Excel to variables for the current iteration
        data.Read
        'Fill catalogues info
        Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/ctxtVIQMFE-FECOD").text = kpi
        Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/txtVIQMFE-FETXT").text = text
        Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/btnRIWO00-IPSDT").press
        Session.findById("wnd[1]/usr/ctxtVIQMFE-BAUTL").text = mrp
        Session.findById("wnd[1]/usr/btnRQM00-KLTEXT").press
        Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = code1
        Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,1]").text = code2
        Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,2]").text = code3
        Sap.Back
        Session.findById("wnd[1]/tbar[0]/btn[6]").press
        'Check if there is an other item to add
        If (objSheet.Cells(i + 1, 1).Value = "") And (objSheet.Cells(i + 1, 7) <> "") Then 'If next serial line is empty and first catalogue cell is not empty add a new item
            Color.Green (i)
            Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/btnRIWO00-INUPS").press
            Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/ctxtVIQMFE-OTEIL").text = "rma"
            i = i + 1
        Else
            flag_next = 0 'Exit the 'add catalogue item' While Loop
        End If
    Wend
End Sub
    
Sub Printer()
    'Print repair log to SAP spooler
    Session.findById("wnd[0]/tbar[0]/btn[86]").press
    Session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS/chkWWORKPAPER-TDIMMED[6,0]").Selected = False
    Session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS").getAbsoluteRow(0).Selected = True
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
End Sub

Sub Status()
    'Append new text to RMA long text if needed
    IW72.set_service_order
    
    'Open user status menu
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
  
    'Press UP arrow if user status is currently at OTV (70)
    If Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/txtANWS_STONR[1,3]").text = "70" Then
        Session.findById("wnd[1]/usr/btnAUP").press
    End If
    
    'Set status TOEV
    If cellB <> "" Then Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[0,0]").Select
    'Set status EVAL
    If cellC <> "" Then Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[1,0]").Select
    'Set status HOLD
    If cellD <> "" Then Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[2,0]").Select
    'Set status REPA
    If cellE <> "" Then Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[3,0]").Select
    'Set status ESCL
    If cellF <> "" Then Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[4,0]").Select
    'Set status OTV
    If cellG <> "" Then
        Session.findById("wnd[1]/usr/btnADOWN").press 'Press menu down arrow
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[4,0]").Select
    End If

    '*********** SUB STATUS SECTION START *********************
    If cellH = "a" Then 'Activate BO status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[0,0]").Selected = True
    ElseIf cellH = "r" Then 'deactivate status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[0,0]").Selected = False
    End If
 
    If cellI = "a" Then 'Activate ENG status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[2,0]").Selected = True
    ElseIf cellI = "r" Then 'deactivate status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[2,0]").Selected = False
    End If

    If cellJ = "a" Then 'Activate FA status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[3,0]").Selected = True
    ElseIf cellJ = "r" Then 'deactivate status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[3,0]").Selected = False
    End If
  
    If cellK = "a" Then 'Activate NPF status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[5,0]").Selected = True
    ElseIf cellK = "r" Then 'deactivate status
        Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[5,0]").Selected = False
    End If
    
    'Check if something is to be done in second half sub status update
    If (cellL & cellM & cellN & cellO & cellP) <> "" Then
        Session.findById("wnd[1]/usr/btnODOWN").press 'Press down on menu
       
        If cellL = "a" Then 'Activate PO status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[0,0]").Selected = True
        ElseIf cellL = "r" Then 'deactivate status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[0,0]").Selected = False
        End If
 
        If cellM = "a" Then 'Activate PRD status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[1,0]").Selected = True
        ElseIf cellM = "r" Then 'deactivate status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[1,0]").Selected = False
        End If
        
        If cellN = "a" Then 'Activate SCRAP status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[2,0]").Selected = True
        ElseIf cellN = "r" Then 'deactivate status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[2,0]").Selected = False
        End If
    
        If cellO = "a" Then 'Activate SWAP status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[3,0]").Selected = True
        ElseIf cellO = "r" Then 'deactivate status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[3,0]").Selected = False
        End If
  
        If cellP = "a" Then 'Activate TS status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[4,0]").Selected = True
        ElseIf cellP = "r" Then 'deactivate status
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-ANWSO[4,0]").Selected = False
        End If
    End If
    
    'Confirm and save user status
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Sap.Save
End Sub

Sub set_service_order()
    If logs <> "" Then
        'Save old repair log text
        text_buffer = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text
        'Write comments+ old buffer
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text = logs + vbCr + text_buffer
    End If
End Sub

Sub Serial_batch()
    data.Serial_batch
    Sap.Go "iw72"
    IW72.Config
    Session.findById("wnd[0]/usr/txtSERIALNR-LOW").text = "blank"
    Session.findById("wnd[0]/usr/btn%_SERIALNR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[16]").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Sap.Execute
    Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    Session.findById("wnd[0]/tbar[1]/btn[37]").press
End Sub

Sub get_stock_status()
    'Go get in_out status
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIOLU").Select
    Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIOLU/ssubSUB_AUFTRAG:SAPLIWOL:0300/tblSAPLIWOLOBJK_120/ctxtRIWOL-SERNR[2,0]").SetFocus
    Session.findById("wnd[0]").sendVKey 2
    sloc = Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122C:SAPLITO0:1220/ctxtEQBS-B_LAGER").text
    If sloc = "PL01" Then
        MsgBox "Board is in PL01, no need for us to make any reservation or movement." + vbCr + "The one wanting the board must create his 901 and 902 in the Plant 1000 Sloc Pl01"
        in_out = "out"
    Else
        in_out = Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122C:SAPLITO0:1220/ctxtEQBS-B_WERK").text
        If in_out = "1010" Then
            in_out = "out"
        Else
            in_out = "in"
        End If
    End If
    Sap.Back
End Sub

Sub get_in_out()
    If sloc <> "PL01" Then
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB").Select
        'Set status
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
        If in_out = "in" Then
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[3,0]").Select
            Session.findById("wnd[1]/tbar[0]/btn[0]").press 'Confirm status menu
            Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4,12]").text = "1"
        Else 'Out
            Session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[4,0]").Select
            Session.findById("wnd[1]/tbar[0]/btn[0]").press 'Confirm status menu
            Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4,12]").text = "-1"
        End If
       
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB").Select
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1,12]").text = part
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8,12]").text = "0001"
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9,12]").text = "1000"
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10,12]").text = "20"
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/btnLTICON-LTOPR[3,12]").press
        
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/btnBTN_MKAG").press
        Session.findById("wnd[0]/tbar[1]/btn[39]").press
        Do While exitflag = 0
            exitflag = 1
            buff_bart = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1300/subSUB_KMP:SAPLCOMD:3001/ctxtRESBD-MATNR").text
            qty = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1300/tabsTS_1300/tabpMKAG/ssubSUB_KMP_DETAIL:SAPLCOMD:3100/txtRESBD-MENGE").text
            If (part <> buff_bart) Or (in_out = "in" And qty = "1-") Then
                Session.findById("wnd[0]/tbar[1]/btn[39]").press
                exitflag = 0
            End If
        Loop
        Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1300/subSUB_KMP:SAPLCOMD:3001/txtRESBD-POTX1").text = serial
        req_num = Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1300/tabsTS_1300/tabpMKAG/ssubSUB_KMP_DETAIL:SAPLCOMD:3100/txtRESBD-RSNUM").text
    End If
End Sub

Sub Full_run()
    IW72.Enter
    IW72.Info
    IW72.get_batch
    IW72.repair_log
    IW72.Printer
End Sub

Sub No_print()
    IW72.Enter
    IW72.Info
    IW72.get_batch
    IW72.repair_log
    Sap.Back
    Sap.Save
End Sub

Sub Read_only()
    IW72.Enter
    IW72.Info
End Sub

Sub req_in_out()
    IW72.Enter
    IW72.Info
    IW72.set_service_order
    IW72.get_stock_status
    IW72.get_in_out
    Sap.Save
End Sub
