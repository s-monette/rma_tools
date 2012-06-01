Attribute VB_Name = "IW42"
Sub teco()
    'Enter IW42
    Sap.Go "iw42"
        
    'Enter service order
    On Error Resume Next 'Cheap fix for SAP erratic way of naming object. Activate error handling
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0201/ctxtCMFUD-AUFNR").text = Order
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0203/ctxtCMFUD-AUFNR").text = Order
    On Error GoTo 0 'Deactivate error handling
    Sap.Enter
    
    'Check if function is currently locked
    While (Session.findById("wnd[0]/sbar").text <> "")
        Sap.Enter
    Wend
    
    'Do TECO and save
    On Error Resume Next 'Cheap fix for SAP erratic way of naming object. Activate error handling
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0201/btnHEADER_TECO").press 'Press TECO button
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0203/btnHEADER_TECO").press
    On Error GoTo 0 'Deactivate error handling
    Sap.Save
End Sub
