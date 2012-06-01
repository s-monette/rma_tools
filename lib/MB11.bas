Attribute VB_Name = "MB11"
Sub Config(ByVal mb11_mvt As String)
    'Enter MB11
    Sap.Go "mb11"
    
    'Setup config
    Session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = mb11_mvt
    Session.findById("wnd[0]/usr/ctxtRM07M-SOBKZ").text = "e"

    Select Case mb11_mvt
        Case "411" 'Mvt from Sales Order to own stock
            Session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "1000"
            Session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "PL01"
        Case "412" 'Mvt from Sales Order to own stock (reversal)
            Session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "1000"
            Session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "PL01"
        Case Else 'Default
            Session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "1010"
            Session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "500"
    End Select
    
    Sap.Enter

    'Input data
    Session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/ctxtMSEGK-MAT_KDAUF").text = rma
    Session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/txtMSEGK-MAT_KDPOS").text = Item
    
    'Destination hardcoded to Plant 1000, Sloc PL01
    Select Case mb11_mvt
        Case "301" ' Mvt from plant to plant
            Session.findById("wnd[0]/usr/ctxtMSEGK-UMWRK").text = "1000"
            Session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").text = "PL01"
        Case "302" ' Mvt from plant to plant (Reversal)
            Session.findById("wnd[0]/usr/ctxtMSEGK-UMWRK").text = "1000"
            Session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").text = "PL01"
    End Select
    
    If (mb11_mvt = "501") Then 'Use partout when the board as been converted.
        Session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = IW72.partout
    Else
        Session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = part
    End If
    
    Session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = "1"
    Sap.Enter
    
    'Check if matarial is EOL, send press enter to confirm if so
    If (Session.findById("wnd[0]/sbar").text <> "") Then Sap.Enter
    
    If batchout <> "" Or batch <> "" Then
        Session.findById("wnd[0]/usr/ctxtMSEG-CHARG").text = batch
        If batchout <> "" Then
            Session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").text = batchout
            Sap.Enter
        Else
            Session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").text = batch
            Sap.Enter
        End If
    End If
    
    Session.findById("wnd[1]/usr/sub:SAPLIPW1:0200/ctxtRIPW0-SERNR[0,2]").text = serial
    
    'Validate and save
    Session.findById("wnd[1]").sendVKey 0
    Sap.Save
End Sub
