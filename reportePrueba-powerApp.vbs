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
 session.findById("wnd[0]").resizeWorkingPane 137,35,false
 session.findById("wnd[0]/tbar[0]/okcd").text = "/niw72"
 session.findById("wnd[0]").sendVKey 0
 session.findById("wnd[0]/usr/chkDY_OFN").selected = false
 session.findById("wnd[0]/usr/chkDY_OFN").setFocus
 session.findById("wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH").press
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select
 session.findById("wnd[1]").sendVKey 4
 session.findById("wnd[2]/usr/lbl[6,16]").setFocus
 session.findById("wnd[2]/usr/lbl[6,16]").caretPosition = 1
 session.findById("wnd[2]").sendVKey 2
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").setFocus
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 0
 session.findById("wnd[1]").sendVKey 4
 session.findById("wnd[2]/usr/lbl[6,18]").setFocus
 session.findById("wnd[2]/usr/lbl[6,18]").caretPosition = 1
 session.findById("wnd[2]").sendVKey 2
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,1]").setFocus
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,1]").caretPosition = 0
 session.findById("wnd[1]").sendVKey 4
 session.findById("wnd[2]/usr/lbl[1,20]").setFocus
 session.findById("wnd[2]/usr/lbl[1,20]").caretPosition = 2
 session.findById("wnd[2]").sendVKey 2
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,1]").setFocus
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,1]").caretPosition = 0
 session.findById("wnd[1]").sendVKey 4
 session.findById("wnd[2]/usr/lbl[1,21]").setFocus
 session.findById("wnd[2]/usr/lbl[1,21]").caretPosition = 2
 session.findById("wnd[2]").sendVKey 2
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,2]").setFocus
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,2]").caretPosition = 0
 session.findById("wnd[1]").sendVKey 4
 session.findById("wnd[2]/usr").verticalScrollbar.position = 1
 session.findById("wnd[2]/usr").verticalScrollbar.position = 2
 session.findById("wnd[2]/usr/lbl[6,28]").setFocus
 session.findById("wnd[2]/usr/lbl[6,28]").caretPosition = 1
 session.findById("wnd[2]").sendVKey 2
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,2]").setFocus
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,2]").caretPosition = 0
 session.findById("wnd[1]").sendVKey 4
 session.findById("wnd[2]/usr").verticalScrollbar.position = 1
 session.findById("wnd[2]/usr").verticalScrollbar.position = 2
 session.findById("wnd[2]/usr/lbl[6,31]").setFocus
 session.findById("wnd[2]/usr/lbl[6,31]").caretPosition = 10
 session.findById("wnd[2]").sendVKey 2
 session.findById("wnd[1]/tbar[0]/btn[8]").press
 session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "1394"
 session.findById("wnd[0]/usr/ctxtPRIOK-LOW").setFocus
 session.findById("wnd[0]/usr/ctxtPRIOK-LOW").caretPosition = 0
 session.findById("wnd[0]/tbar[1]/btn[8]").press
 