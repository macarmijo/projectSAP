'funcion para consumir repuestos - hasta 3 repuestos por orden.
Sub ProcessServiceOrders()
    Dim wsActive As Worksheet
    Dim wsInfo As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim serviceOrder As String, tloc As String
    Dim matNr1 As String, matNr2 As String, matNr3 As String
    Dim batchNr1 As String, batchNr2 As String, batchNr3 As String
    Dim batchCount As Integer

    On Error GoTo ErrorHandler ' Enable error handling

    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    ' If connection is established, continue with the rest of the code
    On Error GoTo 0
    
    ' Set worksheets
    Set wsActive = ThisWorkbook.ActiveSheet
    Set wsInfo = ThisWorkbook.Sheets("info estatica")

    ' Find the last row in column B
    lastRow = wsActive.Cells(wsActive.Rows.Count, "B").End(xlUp).Row
    
    ' Get component numbers from "info estatica" tab
    matNr1 = wsInfo.Cells(6, "E").Value
    matNr2 = wsInfo.Cells(7, "E").Value
    matNr3 = wsInfo.Cells(8, "E").Value

    ' Loop through each service order
    For i = 2 To lastRow
       'On Error GoTo WarningHandler
        serviceOrder = wsActive.Cells(i, "B").Value
        tloc = wsActive.Cells(i, "P").Value

        ' Read batch numbers from columns L, M, N
        batchNr1 = wsActive.Cells(i, "L").Value
        batchNr2 = wsActive.Cells(i, "M").Value
        batchNr3 = wsActive.Cells(i, "N").Value
        
        ' Count non-"NA" and non-empty batch numbers
        batchCount = 0
        If batchNr1 <> "NA" And batchNr1 <> "" Then batchCount = batchCount + 1
        If batchNr2 <> "NA" And batchNr2 <> "" Then batchCount = batchCount + 1
        If batchNr3 <> "NA" And batchNr3 <> "" Then batchCount = batchCount + 1

        ' Debugging information
        Debug.Print "Service Order: " & serviceOrder & ", Batch Count: " & batchCount

        ' Check if batch numbers are different from "NA" and update SAP accordingly
        If batchCount > 0 Then
            ' Start SAP transaction IW32
            session.StartTransaction ("IW32")
            session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = serviceOrder
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpMUEB").Select
            
            ' Update SAP for each batch number
            If batchNr1 <> "NA" And batchNr1 <> "" Then
                Call UpdateSAPComponent(serviceOrder, matNr1, batchNr1, tloc, 1)
            End If
            If batchNr2 <> "NA" And batchNr2 <> "" Then
                Call UpdateSAPComponent(serviceOrder, matNr2, batchNr2, tloc, 2)
            End If
            If batchNr3 <> "NA" And batchNr3 <> "" Then
                Call UpdateSAPComponent(serviceOrder, matNr3, batchNr3, tloc, 3)
            End If
        End If
        
        'Boton Guardar
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        ' Continue to the next row
        'On Error GoTo 0
        'GoTo ContinueLoop
    
'WarningHandler:
        ' Handle the popup warning
        'MsgBox "Popup detected. Please handle it manually and then click OK to continue.", vbExclamation, "SAP Popup Warning"
        'Resume  ' Resume the code after the user handles the popup
    
'ContinueLoop:
    Next i
    
    ' Display a completion message
    MsgBox "Tarea completada.", vbInformation, "Información"
    
    Exit Sub ' Ensure we exit the subroutine before the error handler

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

End Sub

Sub UpdateSAPComponent(serviceOrder As String, materialNumber As String, batchNumber As String, tloc As String, rowIndex As Integer)

    ' Initialize session (example)
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0).Children(0)

    ' Update SAP component
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1," & rowIndex & "]").Text = materialNumber
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4," & rowIndex & "]").Text = "1"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8," & rowIndex & "]").Text = tloc
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9," & rowIndex & "]").Text = "1394"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10," & rowIndex & "]").Text = "0010"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11," & rowIndex & "]").Text = batchNumber
    'session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,1]").SetFocus
    'session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11,1]").caretPosition = 0
    'session.findById("wnd[0]").sendVKey 0

End Sub