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
Dim objSheet, intRow, i
Set objExcel = GetObject(,"Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
For i = 2 to objSheet.UsedRange.Rows.Count

'defino variables que van a obtener sus valores de la hoja de excel activa al momento de correr el script
orderService = Trim(CStr(objSheet.Cells(i, 3).Value))
hsTrabajadas = Trim(CStr(objSheet.Cells(i, 6).Value))
numeroPersona = Trim(CStr(objSheet.Cells(i, 7).Value))
textoDescriptivo = Trim(CStr(objSheet.Cells(i, 14).Value))
equipoSeguridadElectrica = "1000001149900"
'directorio podriamos sacarlo del excel tambien!! VER de modificar!
directorio = "C:\Users\Roldam13\OneDrive - Medtronic PLC\Servicio Tecnico\Comercial\Contratos de Mantenimiento\Hospital Italiano\ABRIL 2024"
archivoPDF = "SE - " + Trim(CStr(objSheet.Cells(i, 5).Value)) + ".pdf"

'entro a la OS mediante iw32
session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = orderService
session.findById("wnd[0]").sendVKey 0
'agrego texto descriptivo del trabajo realizado
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text = "ON DEMAND # PREVENTIVE" + vbCr + "" + vbCr + textoDescriptivo + vbCr + "" + vbCr + ""
'busca la linea donde voy a ingresar horas trabajadas
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").setSelectionIndexes 461,461
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").firstVisibleLine = "3"
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").setSelectionIndexes 0,0
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").firstVisibleLine = "1"
'horas trabajadas
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").text = hsTrabajadas
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").setFocus
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
'ingreso equipo de seguridad electrica en Equipment
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").getAbsoluteRow(0).selected = true
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0,0]").setFocus
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0,0]").caretPosition = 0
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
session.findById("wnd[1]/tbar[0]/btn[5]").press
session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").text = equipoSeguridadElectrica
session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").setFocus
session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").caretPosition = 13
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[29]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
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

'abro noti para subir pdf de SE
session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_ARL_LINK"
session.findById("wnd[1]/usr/txtSCENARIO").setFocus
session.findById("wnd[1]/usr/txtSCENARIO").caretPosition = 17
session.findById("wnd[1]/usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[0]/shell").hierarchyHeaderWidth = 268
session.findById("wnd[1]/usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0000000014","HITLIST"
session.findById("wnd[1]/usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0000000014","HITLIST"
session.findById("wnd[1]/usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[0]/shell").topNode = "0000000002"
'double click PDF attachment
session.findById("wnd[1]/usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0000000014","HITLIST"
session.findById("wnd[2]/usr/txtDY_PATH").text = directorio
session.findById("wnd[2]/usr/txtDY_FILENAME").text = archivoPDF
session.findById("wnd[2]/usr/txtDY_FILENAME").caretPosition = 19
session.findById("wnd[2]").sendVKey 0
session.findById("wnd[2]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").text = "SE"
session.findById("wnd[2]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").caretPosition = 2
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 169,195
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press

'vuelvo a ingresar a la OS
session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0

'dar Complete
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 1
session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 2
session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 3
session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 4
session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 26
session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 25
session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,0]").selected = true
session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
'dar TECO
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").setFocus
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[36]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

next
