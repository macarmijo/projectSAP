Sub completarIW32()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object, campoHorasTrabajadas As Object
    Dim wsActive As Worksheet, wsInfo As Worksheet
    Dim lastRow As Long
    Dim r As Range
    Dim orderService As String, nroSerie As String, hsTrabajadas As String, numeroPersona As String
    Dim textoDescriptivo As String, rigel As String, pts As String, seguridadElectrica As String
    Dim infoCodigo As String, infoCodigo2 As String
    Dim doubleEquip As Boolean
    Dim codigoRigel As String, codigoPTS As String, codigoSE As String

    On Error GoTo ErrorHandler

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)

    Set wsActive = ThisWorkbook.ActiveSheet
    Set wsInfo = ThisWorkbook.Sheets("info estatica")

    codigoRigel = Trim(CStr(wsInfo.Cells(1, 2).Value))
    codigoPTS = Trim(CStr(wsInfo.Cells(2, 2).Value))
    codigoSE = Trim(CStr(wsInfo.Cells(3, 2).Value))

    lastRow = wsActive.Cells(wsActive.Rows.Count, 2).End(xlUp).Row

    For Each r In wsActive.Range(wsActive.Cells(2, 1), wsActive.Cells(lastRow, 1))
        On Error GoTo GeneralErrorHandler

        orderService = Trim(CStr(r.Offset(0, 1).Value))
        Debug.Print "Procesando OS: " & orderService
        textoDescriptivo = Trim(CStr(r.Offset(0, 3).Value))
        hsTrabajadas = Trim(CStr(r.Offset(0, 4).Value))
        rigel = Trim(CStr(r.Offset(0, 5).Value))
        pts = Trim(CStr(r.Offset(0, 6).Value))
        seguridadElectrica = Trim(CStr(r.Offset(0, 7).Value))
        nroSerie = Trim(CStr(r.Offset(0, 10).Value))
        numeroPersona = Trim(CStr(r.Offset(0, 17).Value))

        session.StartTransaction ("IW32")
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        HandleSAPPopups session

        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").Text = _
            "ON DEMAND # PREVENTIVE" & vbCr & vbCr & textoDescriptivo & vbCr & vbCr

        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").Select
        Set campoHorasTrabajadas = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]")
        campoHorasTrabajadas.Text = hsTrabajadas
        campoHorasTrabajadas.SetFocus
        campoHorasTrabajadas.caretPosition = 9
        session.findById("wnd[0]").sendVKey 0

        infoCodigo = ""
        infoCodigo2 = ""
        doubleEquip = False

        Select Case True
            Case LCase(rigel) = "si" And LCase(seguridadElectrica) = "si"
                doubleEquip = True
                infoCodigo = codigoRigel
                infoCodigo2 = codigoSE
            Case LCase(rigel) = "si"
                infoCodigo = codigoRigel
            Case LCase(pts) = "si" And LCase(seguridadElectrica) = "si"
                doubleEquip = True
                infoCodigo = codigoPTS
                infoCodigo2 = codigoSE
            Case LCase(pts) = "si"
                infoCodigo = codigoPTS
            Case LCase(seguridadElectrica) = "si"
                infoCodigo = codigoSE
        End Select

        If doubleEquip Then
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
            session.findById("wnd[1]/tbar[0]/btn[5]").press
            session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[20]").press
            session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo2
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[29]").press
            session.findById("wnd[0]/tbar[0]/btn[3]").press
        ElseIf infoCodigo <> "" Then
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
            session.findById("wnd[1]/tbar[0]/btn[5]").press
            session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[29]").press
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[3]").press
        End If

        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.Position = 22
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").Selected = True
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").SetFocus
        session.findById("wnd[1]/tbar[0]/btn[0]").press

        session.findById("wnd[0]/tbar[0]/btn[11]").press
        HandleSAPPopups session
        r.Offset(0, 17).Value = "Guardado OK"

GoTo ContinueLoop

GeneralErrorHandler:
        r.Offset(0, 17).Value = "Error: " & Err.Description
        Resume ContinueLoop

ContinueLoop:
        On Error GoTo GeneralErrorHandler
    Next r

    MsgBox "Tarea completada.", vbInformation, "Información"
    Set SapGuiAuto = Nothing
    Set SapGuiApp = Nothing
    Set Connection = Nothing
    Set session = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Necesitas abrir SAP.", vbCritical, "Error de Conexión"
End Sub

Sub HandleSAPPopups(ByRef session As Object)
    Dim startTime As Single: startTime = Timer
    Dim popupText As String

    Do While Timer - startTime < 10
        DoEvents
        If session.Children.Count > 1 Then
            On Error Resume Next
            popupText = session.findById("wnd[1]/usr/txtMESSTXT1").Text
            On Error GoTo 0

            If popupText <> "" Then
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                Exit Do
            End If
        End If
    Loop
End Sub