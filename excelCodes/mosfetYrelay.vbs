Sub mosfetYRelay()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim r As Range, c As Range
    Dim orderService As String, numeroPersona As String, mosfetBatch As String, relayBatch As String, tloc As String
    Dim pestanaExcel As Worksheet
    Dim lastRow As Long, componentRelayNum As String, componentMosfetNum As String, plant As String, act As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    ' Set the worksheet
    Set pestanaExcel = ThisWorkbook.ActiveSheet
    
    ' Find the last row with data in column A
    lastRow = pestanaExcel.Cells(pestanaExcel.Rows.Count, "B").End(xlUp).Row
    
    ' Set the component and plant numbers
    componentRelayNum = "230017006"
    componentMosfetNum = "239300044"
    plant = "1394"
    act = "0010" ' Adjust this if it should be different
    
    ' Iterate through each row
    For i = 2 To lastRow
        orderService = Trim(CStr(pestanaExcel.Cells(i, 2).Value))
        mosfetBatch = Trim(CStr(pestanaExcel.Cells(i, 12).Value))
        relayBatch = Trim(CStr(pestanaExcel.Cells(i, 13).Value))
        tloc = Trim(CStr(pestanaExcel.Cells(i, 15).Value))
        numeroPersona = Trim(CStr(pestanaExcel.Cells(i, 16).Value))
        
        ' Start SAP transaction IW32
        session.StartTransaction ("IW32")
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        
        ' Update material numbers in SAP table
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020").Columns.ElementAt(1).Width = 16
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1,0]").Text = componentRelayNum
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1,1]").Text = componentMosfetNum
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4,0]").Text = "1"
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4,1]").Text = "1"
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8,0]").Text = "t008"
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8,1]").Text = "t008"
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9,0]").Text = plant
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9,1]").Text = plant
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10,0]").Text = act
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10,1]").Text = act
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,0]").Text = mosfetBatch
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,1]").Text = relayBatch
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,1]").SetFocus
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,1]").caretPosition = 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        ' Start SAP transaction MIGO_GI
        session.StartTransaction ("migo_gi")
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        
        ' Update the service order in MIGO_GI transaction
        'session.findById("wnd[0]/usr/subSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDATAX:SAPLMIGO:1505/tblSAPLMIGOTC_CUSTOM/ctxtGOITEM-RESERV_NO[11,0]").Text = orderService
        'session.findById("wnd[0]/usr/subSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDATAX:SAPLMIGO:1505/tblSAPLMIGOTC_CUSTOM/ctxtGOITEM-RES_ITEM[12,0]").Text = "1"
        'session.findById("wnd[0]/usr/subSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDATAX:SAPLMIGO:1505/tblSAPLMIGOTC_CUSTOM/ctxtGOITEM-BATCH[16,0]").Text = mosfetBatch
        'session.findById("wnd[0]").sendVKey 0
        'session.findById("wnd[0]/usr/txtMKPF-BKTXT").Text = numeroPersona
        'session.findById("wnd[0]/tbar[0]/btn[11]").press
        'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    Next i
    
   
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
    
End Sub
