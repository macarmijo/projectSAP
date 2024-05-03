'Script para completar las ordenes de field Action. Agrega texto, consume horas y cierra ZC. 

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
 
'Conexion con hoja de excel activa de donde saco y guardo informacion
Dim objExcel
Dim intRow, i
Set objExcel = GetObject(,"Excel.Application")
Set objWorkbook = objExcel.ActiveWorkbook
Set sheet1 = objWorkbook.Worksheets(1)

horas = Trim(CStr(sheet1.Cells(2, 9).Value))
nroPersona = Trim(CStr(sheet1.Cells(2, 10).Value))

'arranca en fila 2 - fila 1 tiene nombre de los valores de la columna
For i = 2 to sheet1.UsedRange.Rows.Count
    notif = Trim(CStr(sheet1.Cells(i, 1).Value))
    orderText = Trim(CStr(sheet1.Cells(i, 8).Value))
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = notif
    session.findById("wnd[0]").sendVKey 0
    
    'guarda en excel el nro de OS
    sheet1.Cells(i, 2).Value = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text
    session.findById("wnd[0]").sendVKey 0
    os = Trim(CStr(sheet1.Cells(i, 2).Value))
    
    'entrar OS
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = os
    session.findById("wnd[0]").sendVKey 0

    'texto
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text = "FIELD ACTION # FIELD ACTION" + vbCr + "" + vbCr + orderText + vbCr + ""

    'pongo horas trabajadas
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").setSelectionIndexes 181,181
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").setSelectionIndexes 0,0
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").text = horas
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").setFocus
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]").caretPosition = 9
    session.findById("wnd[0]").sendVKey 0

    'guardar cambios
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    'consumo horas trabajadas
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw42"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").text = nroPersona
    session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").setFocus
    session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").caretPosition = 7
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    'vuelvo a ingresar a la OS
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    'dar Complete
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 3
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 6
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 9
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 12
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 15
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 18
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 21
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,4]").selected = true
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,4]").setFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    'dar TECO
    session.findById("wnd[0]/tbar[1]/btn[36]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
 
 next