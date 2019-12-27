Set WshShell = WScript.CreateObject("WScript.Shell")

	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", "", "", "open", 1
	
WScript.Sleep 500 
	
	If Not IsObject(app) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set app = SapGuiAuto.GetScriptingEngine
    End If
	
	If Not IsObject(Connection) Then
        Set Connection = app.OpenConnectionByConnectionString("StringDeConexão", True)
    End If
	
	If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
        
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject app, "on"
    End If
	
		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/tbar[0]/okcd").Text = "iw53"
		session.findById("wnd[0]").sendVKey 0
		
		session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = "319028776"
		
		session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").caretPosition = 9
		session.findById("wnd[0]").sendVKey 0
		
		session.findById("wnd[0]/shellcont/shell").selectItem "010", "Column01"
		session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem "010", "Column01"
		session.findById("wnd[0]/shellcont/shell").clickLink "010", "Column01"
        
