dim currentPassword
dim newPassword
currentPassword = ""
newPassword = ""
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 128,39,false
session.findById("wnd[0]/tbar[0]/okcd").text = "su3"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[6]").press
session.findById("wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-BCODE").text = currentPassword
session.findById("wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-NCODE").text = newPassword
session.findById("wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-NCOD2").text = newPassword
session.findById("wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-NCOD2").setFocus
session.findById("wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-NCOD2").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[0]").press
