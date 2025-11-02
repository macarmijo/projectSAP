'codigo probado 4/06 funciona . PERO si ya tiene un equipment cargado, nmo lo pisa, corta en script, no guarda la orden y deja en error. Hay que realizar un IF
'SIIIII ya hay contenido en linea 1, entonces no adjuntar equipment.... revisar!

'11/06 agregue lo de equipment si ya hay algo que lo saltee. Agregue acty tye pero rompe el codigo asi que esta comentado
'compl quedo harcodeado porque de otra forma no encontre solucion aun
'todo el resto funciona muy bien. ver de pasar a funciones para que no sea tan pesado el codigo
'falta chequeo de nro de persona

'14/09 agrege codigo de repuestos y mas errores porque el ejemplo de FLOSAN tiene customer block al inicio y popups al guardar
'uno que con enter se va normal, otro de cost error que hay qye darle al boton de YES para que se vaya. (agregado en HandleSapPopups)
'funciona super :)
'-> falta: control pesona responsable en iw32 y iw52. chequeo de PRTP sino dar bandera verde.

Sub IW32()
    Dim SapGuiAuto As Object, SapGuiApp As Object, connection As Object, session As Object, campoHorasTrabajadas As Object
    Dim wsActive As Worksheet, wsInfo As Worksheet
    Dim r As Range
    Dim orderService As String, nroSerie As String, hsTrabajadas As String, numeroPersona As String
    Dim textoDescriptivo As String, rigel As String, pts As String, seguridadElectrica As String
    Dim infoCodigo As String, infoCodigo2 As String
    Dim doubleEquip As Boolean
    Dim codigoRigel As String, codigoPTS As String, codigoSE As String
    Dim row As Long: row = 3
    Dim sbarText, sbarTextEquipment, actType, fieldValue As String, sloc As String, sysStatus As String
    Dim material(1 To 10) As String
    Dim batch(1 To 10) As String
    Dim j As Integer, rowIndex As Integer
    Dim batchCount As Integer
    Dim rowExcel As Long

    ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    MsgBox "Presione OK si desea completar las OS (IW32)"

    Set wsActive = ThisWorkbook.Sheets("bot-hyperautomate")
    Set wsInfo = ThisWorkbook.Sheets("datos")

    codigoSE = Trim(CStr(wsActive.Cells(1, 4).Value))
    codigoRigel = Trim(CStr(wsActive.Cells(1, 5).Value))
    codigoPTS = Trim(CStr(wsActive.Cells(1, 6).Value))
    actType = Trim(CStr(wsInfo.Cells(3, 10).Value))

    application.EnableCancelKey = xlErrorHandler

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
        tloc = Trim(CStr(r.Offset(0, 9).Value))               ' Columna I
        numeroPersona = Trim(CStr(r.Offset(0, 11).Value))     ' Columna J

        '––– Iniciar IW32 –––
        session.StartTransaction "IW32"
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        HandleSAPPopups session, r
        
        '--- Chequeo bandera verde ---
        sysStatus = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT").Text
        If InStr(1, LCase(sysStatus), "CTRD") > 0 Then
            session.findById("wnd[0]/tbar[1]/btn[25]").press
        End If

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
           session.findById("wnd[0]").sendVKey 0
           HandleSAPPopups session, r
        End If
        
        ' Verificar si hay "ok" en alguna de las variables antes de entrar a equipment
        If LCase(rigel) = "ok" Or LCase(seguridadElectrica) = "ok" Or LCase(pts) = "ok" Then
            'entrar a equipment
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/" & _
                        "ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
        
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
           
        ' Detectar si el popup de nuevo equipment está abierto
        popupExists = False
        On Error Resume Next
        popupExists = Not session.findById("wnd[1]/usr/txtAFFHD-PSNFH") Is Nothing
        On Error GoTo 0

        If popupExists Then
            ' Completar datos en popup
            If doubleEquip Then
                session.findById("wnd[1]/tbar[0]/btn[5]").press
                session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
                session.findById("wnd[1]").sendVKey 0
                session.findById("wnd[1]/tbar[0]/btn[20]").press
                session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo2
                session.findById("wnd[1]").sendVKey 0
                session.findById("wnd[1]/tbar[0]/btn[29]").press
                session.findById("wnd[0]/tbar[0]/btn[3]").press
            ElseIf infoCodigo <> "" Then
                session.findById("wnd[1]/tbar[0]/btn[5]").press
                session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
                session.findById("wnd[1]").sendVKey 0
                session.findById("wnd[1]/tbar[0]/btn[29]").press
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/tbar[0]/btn[3]").press
            End If
        Else
            ' Ya hay equipos cargados, no agregar
            session.findById("wnd[0]/tbar[0]/btn[3]").press
        End If
        End If
        
        'repuestos
        Call cargarRepuestos(session, row)
        
        ' COMPL - o CMPL QTTM - ver de agregar opcion
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/" & _
            "subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.Position = 23
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").Selected = True
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        ' Boton Guardar
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        HandleSAPPopups session, r
        
        ' Capturar texto de barra estado
        On Error Resume Next
            sbarText = session.findById("wnd[0]/sbar").Text
        On Error GoTo 0

        ' Validar resultado guardado
        If InStr(1, LCase(sbarText), "saved with notification") > 0 Then
            r.Offset(0, 21).Value = "Guardado OK"
            r.Offset(0, 1).Interior.Color = RGB(198, 239, 206) ' Verde
        ElseIf InStr(1, LCase(sbarText), "person is not qualified") > 0 Then
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
        Else
            r.Offset(0, 21).Value = "Error al guardar: " & sbarText
            r.Offset(0, 1).Interior.Color = RGB(255, 199, 206) ' Rojo
        End If
        
        GoTo ContinueLoop

GeneralErrorHandler:
        r.Offset(0, 23).Value = "Error: " & Err.Description
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
    Set connection = Nothing
    Set session = Nothing
    Exit Sub

End Sub
Sub HandleSAPPopups(ByRef session As Object, ByRef r As Range)
    Dim popupText As String
    Dim sbarType As String, sbarText As String
    Dim mensajes As String
    Dim maxIter As Integer, iter As Integer
    
    mensajes = ""
    maxIter = 5   ' seguridad: máximo 5 iteraciones
    iter = 0
    
    Do
        iter = iter + 1
        On Error Resume Next
    
        '––– 1. POPUPS –––
        If session.Children.Count > 1 Then
            popupText = session.findById("wnd[1]/usr/txtMESSTXT1").Text
            Debug.Print "Popup detectado: " & popupText
    
            Select Case LCase(Trim(session.findById("wnd[1]").Text))
                Case "cost calculation"
                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                Case Else
                    session.findById("wnd[1]").sendVKey 0   ' ENTER genérico
            End Select
        End If
        
        '––– 2. BARRA DE ESTADO –––
        sbarType = ""
        sbarText = ""
        On Error Resume Next
        sbarType = session.findById("wnd[0]/sbar").MessageType
        sbarText = session.findById("wnd[0]/sbar").Text
        On Error GoTo 0
        
        If sbarText <> "" Then
            mensajes = mensajes & "Barra (" & sbarType & "): " & sbarText & vbCrLf
            
            Select Case sbarType
                Case "W" ' Warning ? ENTER y continuar
                    session.findById("wnd[0]").sendVKey 0
                Case "E" ' Error ? registrar y salir de la orden
                    r.Offset(0, 23).Value = "Error: " & sbarText
                    r.Offset(0, 0).Interior.Color = RGB(255, 199, 206) ' rojo
                    session.findById("wnd[0]").sendVKey 15   ' back
                    Exit Sub
            End Select
        End If
        
        ' salir si no hay popup ni mensaje
        If session.Children.Count = 1 And sbarText = "" Then Exit Do
        DoEvents
    Loop While iter < maxIter
    
    ' Guardar los mensajes (si hubo) en la columna V
    If mensajes <> "" Then
        r.Offset(0, 24).Value = Trim(mensajes)
    End If
End Sub

Sub cargarRepuestos(ByRef session As Object, ByVal rowExcel As Long)

    Dim wsActive As Worksheet, wsDatos As Worksheet
    Dim material As String, qty As String, batch As String
    Dim sloc As String, planta As String
    Dim i As Integer, rowIndex As Integer
    Dim baseId As String
    
    Set wsActive = ThisWorkbook.Sheets("bot-hyperautomate")
    Set wsDatos = ThisWorkbook.Sheets("datos")
    
    sloc = Trim(CStr(wsActive.Cells(rowExcel, 9).Value))    ' Columna I
    planta = Trim(CStr(wsActive.Cells(rowExcel, 10).Value)) ' Columna J
    
    ' ?? Abrir pestaña de componentes (MUEB)
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/" & _
                     "ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB").Select
    
    baseId = "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/" & _
             "ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/" & _
             "ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/"
    
    rowIndex = 0
    
    For i = 6 To 15
    material = Trim(CStr(wsDatos.Cells(i, 12).Value))   ' Col L
    qty = Trim(CStr(wsDatos.Cells(i, 13).Value))        ' Col M
    batch = Trim(CStr(wsActive.Cells(rowExcel, 11 + (i - 5)).Value)) ' Col L–U
    
    Debug.Print "Row SAP " & rowIndex & " | Material=" & material & " | Qty=" & qty & " | Batch=" & batch
    
    If qty <> "" And (batch <> "" And UCase(batch) <> "NA") Then
        ' Verificar si ya existe material cargado en esa fila
        Dim existingMat As String
        existingMat = Trim(session.findById(baseId & "ctxtRESBD-MATNR[1," & rowIndex & "]").Text)
        
        If existingMat <> "" Then
            Debug.Print "Fila " & rowIndex & " ya tiene material (" & existingMat & "), se saltea."
        Else
            ' Cargar nuevo material
            session.findById(baseId & "ctxtRESBD-MATNR[1," & rowIndex & "]").Text = material
            session.findById(baseId & "txtRESBD-MENGE[4," & rowIndex & "]").Text = qty
            session.findById(baseId & "ctxtRESBD-LGORT[8," & rowIndex & "]").Text = sloc
            session.findById(baseId & "ctxtRESBD-WERKS[9," & rowIndex & "]").Text = planta
            session.findById(baseId & "txtRESBD-VORNR[10," & rowIndex & "]").Text = "0010"
            session.findById(baseId & "ctxtRESBD-CHARG[11," & rowIndex & "]").Text = batch
            Debug.Print ">>> Repuesto cargado en SAP fila " & rowIndex
        End If
        
        rowIndex = rowIndex + 1
    End If
Next i

End Sub

'created by Maca Armijo
