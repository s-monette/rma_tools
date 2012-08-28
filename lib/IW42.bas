Attribute VB_Name = "IW42"
Sub teco()
    'Enter IW42
    Sap.Go "iw42"
        
    'Enter service order
    On Error Resume Next 'Cheap fix for SAP erratic way of naming object. Deactivate error handling
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0201/ctxtCMFUD-AUFNR").text = Order
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0203/ctxtCMFUD-AUFNR").text = Order
    On Error GoTo 0 'Reactivate error handling
    Sap.Enter
    
    'Check if function is currently locked
    While (Session.findById("wnd[0]/sbar").text <> "")
        Sap.Enter
    Wend

    'Press TECO button
    On Error Resume Next 'Cheap fix for SAP erratic way of naming object. Deactivate error handling
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0201/btnHEADER_TECO").press
        Session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0203/btnHEADER_TECO").press
        
        'Have to check content of the pop-up box with safety off.
        err_sap = Session.findById("/app/con[0]/ses[0]/wnd[1]/usr/txtMESSTXT2").text
    On Error GoTo 0 'Reactivate error handling
    
    'Second error trap if someone have a service order open, loop if then.
    If err_sap = "cannot be adjusted" Then
        MsgBox "Find who is blocking service order"
        IW42.teco
    End If
        
    Sap.Save
End Sub
