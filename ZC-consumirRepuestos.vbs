'Este codigo realiza los siguientes pasos
'Entra a la orden de servicio iw32 - pega texto descriptivo del trabajo - pone las horas trabajadas - pone el equipo de seguridad electrica
'Luego se consumen las horas trabajadas con iw42
'Se entra a la Noti iw52 para adjuntar el pdf de seguridad electrica segun SE - nro de serie
'Finalmente vuelve a entrar a la OS iw32 para dar CMPL y TECO.
'Fin del proceso

'No borrar - config de SAP
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
 
'esto capaz es innecesario, solo maximiza la pantalla de sap cuando corre el script
session.findById("wnd[0]").maximize
 
'Conexion con hoja de excel activa de donde saco y guardo informacion
Dim objExcel
Dim intRow, i
Set objExcel = GetObject(,"Excel.Application")
Set objWorkbook = objExcel.ActiveWorkbook
Set sheet1 = objWorkbook.Worksheets(1)
Set sheet2 = objWorkbook.Worksheets(2)
Set sheet3 = objWorkbook.Worksheets(3)
Set sheet4 = objWorkbook.Worksheets(4)
Set sheet5 = objWorkbook.Worksheets(5)

numeroPersona = Trim(CStr(sheet1.Cells(3, 2).Value))
hsTrabajadas = Trim(CStr(sheet1.Cells(7, 2).Value))
hose = "10088303"
nozzle = "10088303"

For i = 2 to sheet2.UsedRange.Rows.Count
    notif = Trim(CStr(sheet2.Cells(i, 1).Value))
    textoDescriptivo = Trim(CStr(sheet2.Cells(i, 5).Value))
    batchHose = Trim(CStr(sheet4.Cells(i, 3).Value))
    batchNozzle = Trim(CStr(sheet5.Cells(i, 3).Value))

    'entro a la noti
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = notif
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0

    'guarda en excel el nro de OS
    sheet2.Cells(i, 2).Value = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text
    session.findById("wnd[0]").sendVKey 0
    orderService = Trim(CStr(sheet2.Cells(i, 2).Value))

    'entro a la OS segun el nro que guarde en el paso anterior
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = orderService
    session.findById("wnd[0]").sendVKey 0

    'agrego texto descriptivo del trabajo realizado
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text = "ON DEMAND # PREVENTIVE" + vbCr + "" + vbCr + textoDescriptivo + vbCr + "" + vbCr + ""
    
    'pongo horas trabajadas
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").setSelectionIndexes 146,146
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").setSelectionIndexes 0,0
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").text = hsTrabajadas
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").setFocus
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").caretPosition = 9
    session.findById("wnd[0]").sendVKey 0 
    
    'pongo que repuestos use con su batch tloc y demas
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB").select
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1,0]").text = hose
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4,0]").text = "1"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8,0]").text = "t008"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9,0]").text = "1394"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10,0]").text = "0010"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,0]").text = batchHose
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-WEMPF[13,0]").setFocus
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-WEMPF[13,0]").caretPosition = 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    'consumo horas trabajadas
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw42"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").text = numeroPersona
    session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").setFocus
    session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").caretPosition = 7
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    'consumo repuestos
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmigo_gi"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").topNode = "          1"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = orderService
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").caretPosition = 9
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").text = batchHose
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").setFocus
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").caretPosition = 10
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = true
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0,0]").setFocus
   session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0,0]").press
   session.findById("wnd[0]/tbar[1]/btn[7]").press
   session.findById("wnd[0]/tbar[1]/btn[23]").press

    'vuelvo a ingresar a la OS
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0

   'dar Complete
   session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 19
   session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 20
   session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 21
   session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 22
   session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").selected = true
   session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").setFocus
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   'dar TECO
   session.findById("wnd[0]/tbar[1]/btn[36]").press
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[1]/tbar[0]/btn[0]").press
 
next
