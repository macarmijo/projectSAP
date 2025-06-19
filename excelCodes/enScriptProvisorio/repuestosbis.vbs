
Sub consumirRepuestos()
    Dim wsActive As Worksheet
    Dim wsInfo As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim serviceOrder As String, tloc As String
    Dim material(1 To 3) As String
    Dim batch(1 To 3) As String
    Dim j As Integer, rowIndex As Integer
    Dim batchCount As Integer

    On Error GoTo ErrorHandler

    ' SAP Connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)

    Set wsActive = ThisWorkbook.ActiveSheet
    Set wsInfo = ThisWorkbook.Sheets("info estatica")
    lastRow = wsActive.Cells(wsActive.Rows.Count, "B").End(xlUp).Row

    ' Materiales fijos definidos por ingeniero
    material(1) = CStr(wsInfo.Cells(6, "E").Value) ' Rep 1
    material(2) = CStr(wsInfo.Cells(7, "E").Value) ' Rep 2
    material(3) = CStr(wsInfo.Cells(8, "E").Value) ' Rep 3
    tloc = CStr(wsActive.Cells(2, "P").Value)

    ' Loop por cada orden
    For i = 2 To lastRow
        serviceOrder = CStr(wsActive.Cells(i, "B").Value)
        batch(1) = CStr(wsActive.Cells(i, "L").Value)
        batch(2) = CStr(wsActive.Cells(i, "M").Value)
        batch(3) = CStr(wsActive.Cells(i, "N").Value)

        ' Calcular cuántos repuestos válidos hay
        batchCount = 0
        For j = 1 To 3
            If batch(j) <> "NA" And batch(j) <> "" Then
                batchCount = batchCount + 1
            End If
        Next j

        If batchCount > 0 Then
            ' Ingresar a IW32
            session.StartTransaction ("IW32")
            session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = serviceOrder
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpMUEB").Select

            rowIndex = 0
            For j = 1 To 3
                If batch(j) <> "NA" And batch(j) <> "" Then
                    Call UpdateSAPComponent(serviceOrder, material(j), batch(j), tloc, rowIndex)
                    rowIndex = rowIndex + 1
                End If
            Next j

            ' Guardar en IW32
            session.findById("wnd[0]/tbar[0]/btn[11]").press

            ' Aceptar popups si aparecen
            If session.Children.Count > 1 Then
                session.findById("wnd[0]").sendVKey 0
            End If

            Debug.Print "Pasando batchCount como: " & batchCount & " (" & TypeName(batchCount) & ")"
            Call ejecutarMIGO_GI_dinamico(serviceOrder, material, batch, batchCount)
        End If
    Next i

    MsgBox "Tarea completada.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub

Sub UpdateSAPComponent(serviceOrder As String, materialNumber As String, batchNumber As String, tloc As String, rowIndex As Integer)
    ' Inicializar sesión
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0).Children(0)

    ' Update SAP component
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1," & rowIndex & "]").Text = materialNumber
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4," & rowIndex & "]").Text = "1"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8," & rowIndex & "]").Text = tloc
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9," & rowIndex & "]").Text = "1394"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10," & rowIndex & "]").Text = "0010"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-CHARG[11," & rowIndex & "]").Text = batchNumber
    

End Sub

Sub ejecutarMIGO_GI_dinamico(serviceOrder As String, material() As String, batch() As String, batchCount As Integer)
    Dim session As Object
    Dim i As Integer
    Dim serialInputID As String
    Dim batchFieldID As String
    Dim batchValue As String
    Dim serialField As Object

    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0).Children(0)

    ' Abrir MIGO GI
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmigo_gi"
    session.findById("wnd[0]").sendVKey 0

    ' Ingresar orden
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/" & _
                     "subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = serviceOrder
    session.findById("wnd[0]").sendVKey 0

    ' Marcar TAKE y refrescar
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/" & _
                     "subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").Selected = True
    session.findById("wnd[0]").sendVKey 2

    If batchCount = 1 Then
        i = 0 ' No es necesario hacer clic
    Else
        For i = 0 To batchCount - 1
            ' Hacer clic en el número de línea para expandir
            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/" & _
                             "subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0," & i & "]").SetFocus
            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/" & _
                             "subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0," & i & "]").press

            ' Ir a pestaña Serial Number
            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/" & _
                             "subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/" & _
                             "tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").Select

            serialInputID = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/" & _
                            "subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/" & _
                            "tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/" & _
                            "tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]"

            On Error Resume Next
            Set serialField = session.findById(serialInputID, False)
            On Error GoTo 0

            If Not serialField Is Nothing Then
                If Trim(serialField.Text) = "" Then
                    ' Obtener batch desde la línea correspondiente
                    batchFieldID = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/" & _
                                   "subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-CHARG[10," & i & "]"
                    batchValue = session.findById(batchFieldID).Text
                    serialField.Text = batchValue
                End If
            End If
        Next i
    End If

    ' Finalizar
    session.findById("wnd[0]/tbar[1]/btn[7]").press ' Check
    If session.Children.Count > 1 Then session.Children(1).sendVKey 0 ' Aceptar popup si aparece
    session.findById("wnd[0]/tbar[1]/btn[23]").press ' Post
End Sub