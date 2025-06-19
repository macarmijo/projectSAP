'codigo probado 4/06 funciona . PERO si ya tiene un equipment cargado, nmo lo pisa, corta en script, no guarda la orden y deja en error. Hay que realizar un IF
'SIIIII ya hay contenido en linea 1, entonces no adjuntar equipment.... revisar!

Sub IW32()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object, campoHorasTrabajadas As Object
    Dim wsActive As Worksheet, wsInfo As Worksheet
    Dim r As Range
    Dim orderService As String, nroSerie As String, hsTrabajadas As String, numeroPersona As String
    Dim textoDescriptivo As String, rigel As String, pts As String, seguridadElectrica As String
    Dim infoCodigo As String, infoCodigo2 As String
    Dim doubleEquip As Boolean
    Dim codigoRigel As String, codigoPTS As String, codigoSE As String
    Dim row As Long: row = 3
    Dim sbarText As String
    Dim material(1 To 10) As String
    Dim batch(1 To 10) As String
    Dim j As Integer, rowIndex As Integer
    Dim batchCount As Integer

    On Error GoTo ErrorHandler

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)

    Set wsActive = ThisWorkbook.ActiveSheet
    Set wsInfo = ThisWorkbook.Sheets("datos")

    codigoSE = Trim(CStr(wsActive.Cells(1, 4).Value))
    codigoRigel = Trim(CStr(wsActive.Cells(1, 5).Value))
    codigoPTS = Trim(CStr(wsActive.Cells(1, 6).Value))
    'Debug.Print codigoSE

    Application.EnableCancelKey = xlErrorHandler

    Do While Not UCase(Trim(wsActive.Cells(row, 2).Value)) = "END"
        If Err.Number = 18 Then
            MsgBox "Proceso cancelado por el usuario.", vbExclamation
            Exit Do
        End If
        DoEvents
        Set r = wsActive.Cells(row, 1)
        On Error GoTo GeneralErrorHandler

        orderService = Trim(CStr(r.Offset(0, 1).Value))       ' Columna B
        nroSerie = Trim(CStr(r.Offset(0, 2).Value))           ' Columna C
        seguridadElectrica = Trim(CStr(r.Offset(0, 3).Value)) ' Columna D
        rigel = Trim(CStr(r.Offset(0, 4).Value))              ' Columna E
        pts = Trim(CStr(r.Offset(0, 5).Value))                ' Columna F
        textoDescriptivo = Trim(CStr(r.Offset(0, 6).Value))   ' Columna G
        hsTrabajadas = Trim(CStr(r.Offset(0, 7).Value))       ' Columna H
        tloc = Trim(CStr(r.Offset(0, 8).Value))               ' Columna I
        numeroPersona = Trim(CStr(r.Offset(0, 9).Value))      ' Columna J

        '––– Iniciar IW32 –––
        session.StartTransaction "IW32"
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        HandleSAPPopups session

        '––– Texto descriptivo –––
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").Text = _
            "ON DEMAND # PREVENTIVE" & vbCr & vbCr & textoDescriptivo & vbCr & vbCr

        '––– Horas trabajadas –––
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").Select
        Set campoHorasTrabajadas = session.findById( _
            "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/" & _
            "ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ARBEI[10,0]")
        campoHorasTrabajadas.Text = hsTrabajadas
        campoHorasTrabajadas.SetFocus
        campoHorasTrabajadas.caretPosition = 9
        session.findById("wnd[0]").sendVKey 0

        'verificar actType
        fieldValue = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-LARNT[16,0]").Text
        
        If (Trim(fieldValue) = "") Then
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-LARNT[16,0]").Text = actType
        End If

        ' Determinar códigos de equipo
        infoCodigo = ""
        infoCodigo2 = ""
        doubleEquip = False
        Select Case True
            Case LCase(rigel) = "ok" And LCase(seguridadElectrica) = "ok"
                doubleEquip = True: infoCodigo = codigoRigel: infoCodigo2 = codigoSE
            Case LCase(rigel) = "ok"
                infoCodigo = codigoRigel
            Case LCase(pts) = "ok" And LCase(seguridadElectrica) = "ok"
                doubleEquip = True: infoCodigo = codigoPTS: infoCodigo2 = codigoSE
            Case LCase(pts) = "ok"
                infoCodigo = codigoPTS
            Case LCase(seguridadElectrica) = "ok"
                infoCodigo = codigoSE
        End Select

        ' Ingresar equipment - verificar si ya tiene cargados los equipment entonces pasar al siguiente paso
        If doubleEquip Then
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/" & _
                "ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
            session.findById("wnd[1]/tbar[0]/btn[5]").press
            session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[20]").press
            session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo2
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[29]").press
            session.findById("wnd[0]/tbar[0]/btn[3]").press
        ElseIf infoCodigo <> "" Then
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/" & _
                "ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
            session.findById("wnd[1]/tbar[0]/btn[5]").press
            session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[29]").press
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[3]").press
        End If

        ' COMPL - ver de agregar opcion o CMPL QTTM -
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/" & _
            "subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.Position = 22
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").Selected = True
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        ' Boton Guardar
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        ' 2) Capturamos el texto de la barra de estado
        On Error Resume Next
        sbarText = session.findById("wnd[0]/sbar").Text
        On Error GoTo 0

        ' 3) Si vemos el mensaje de éxito...
        If InStr(1, LCase(sbarText), "saved with notification") > 0 Then
            r.Offset(0, 21).Value = "Guardado OK"
            r.Offset(0, 1).Interior.Color = RGB(198, 239, 206) ' Verde

        ' 4) Si vemos el warning de persona no calificada...
        ElseIf InStr(1, LCase(sbarText), "person is not qualified") > 0 Then
            ' 4a)ENTER
            session.findById("wnd[0]").sendVKey 0
            On Error Resume Next
            sbarText = session.findById("wnd[0]/sbar").Text
            On Error GoTo 0

            If InStr(1, LCase(sbarText), "saved with notification") > 0 Then
                r.Offset(0, 21).Value = "Guardado OK"
                r.Offset(0, 1).Interior.Color = RGB(198, 239, 206)
            Else
                r.Offset(0, 21).Value = "Error al guardar: " & sbarText
                r.Offset(0, 1).Interior.Color = RGB(255, 199, 206) ' Rojo
            End If

        ' 5) Cualquier otro texto distinto lo consideramos fallo
        Else
            r.Offset(0, 21).Value = "Error al guardar: " & sbarText
            r.Offset(0, 1).Interior.Color = RGB(255, 199, 206)
        End If
        
        
        GoTo ContinueLoop

GeneralErrorHandler:
        r.Offset(0, 21).Value = "Error: " & Err.Description
        r.Offset(0, 1).Interior.Color = RGB(255, 199, 206) ' Rojo
        Resume ContinueLoop

ContinueLoop:
        On Error GoTo GeneralErrorHandler
        row = row + 1
    Loop
    
    session.StartTransaction "IW32"
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