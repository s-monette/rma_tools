Attribute VB_Name = "Sap"
Public Session As Object

Sub hook()
On Error GoTo ErrorHandler 'Start SAP and login if it's not already running
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
    Set SapConnection = SapApp.Children(0)
    Set Session = SapConnection.Children(0)
Exit Sub
ErrorHandler:
    Call Sap.Login
End Sub

Sub Login()
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe")

    Do Until Success = True
        Success = wshShell.AppActivate("sap") 'Buggy if other app with "sap" in their name are open
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
    Set SapConnection = SapApp.Openconnection("20 - IBM - Global PSK", True)
    Set Session = SapConnection.Children(0)
    
    Do Until Session.findById("wnd[0]/titl").text = "SAP Easy Access"
        Application.Wait Now + TimeValue("00:00:03")
    Loop

    Sap.hook
End Sub

Sub Save()
    Session.findById("wnd[0]/tbar[0]/btn[11]").press
End Sub

Sub Back()
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub

Sub Enter()
    Session.findById("wnd[0]").sendVKey 0
End Sub

Sub Execute()
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
End Sub

Sub Go(data As String)
    Session.findById("wnd[0]/tbar[0]/okcd").text = "/n" & data
    Sap.Enter
End Sub
