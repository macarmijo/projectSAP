'Se logro bandera verde con t02 y ademas tmb se logro cliente bloqueado
'carga equipment bien. 1 o 2 . queda chequear mejor el error si esta pasado de la fecha de calibracion PERO no deberia tener este error porque NO deberiuamos usar equipos fuera de fecha de calibracion
'se agrega corte para que el usuario pueda elegir desde que LINEA de excel quiere correr el script

Sub IW32()
    Dim SapGuiAuto As Object, SapGuiApp As Object, connection As Object, session As Object, campoHorasTrabajadas As Object
    Dim wsActive As Worksheet, wsInfo As Worksheet
    Dim r As Range
    Dim orderService As String, nroSerie As String, hsTrabajadas As String
    Dim textoDescriptivo As String, rigel As String, pts As String, seguridadElectrica As String
    Dim infoCodigo As String, infoCodigo2 As String
    Dim doubleEquip As Boolean
    Dim codigoRigel As String, codigoPTS As String, codigoSE As String
    Dim row As Long
    Dim sbarText As String, sbarTextEquipment As String, actType As String, fieldValue As String, sloc As String, sysStatus As String
    Dim material(1 To 10) As String
    Dim batch(1 To 10) As String
    Dim j As Integer, rowIndex As Integer
    Dim batchCount As Integer
    Dim rowExcel As Long
    Dim repuestos As String
    Dim popupExists As Boolean

    ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    MsgBox "Presione OK si desea completar las OS (IW32)"

    Set wsActive = ThisWorkbook.Sheets("MP")
    Set wsInfo = ThisWorkbook.Sheets("datos")

    codigoSE = Trim(CStr(wsInfo.Cells(6, "B").Value))
    codigoRigel = Trim(CStr(wsInfo.Cells(6, "C").Value))
    codigoPTS = Trim(CStr(wsInfo.Cells(6, "D").Value))
    actType = Trim(CStr(wsInfo.Cells(3, 10).Value))

    application.EnableCancelKey = xlErrorHandler

    row = Trim(CStr(wsActive.Cells(1, "O").Value))
    If (row < 3) Then
        MsgBox "Debe corregir el inicio del sciprt! No puede ser un numero menor a 3", vbCritical
    End If

    Do While Not UCase(Trim(wsActive.Cells(row, 3).Value)) = "END"
        If Err.Number = 18 Then
            MsgBox "Proceso cancelado por el usuario.", vbExclamation
            Exit Do
        End If
        
        DoEvents
        Set r = wsActive.Cells(row, 1)
        On Error GoTo GeneralErrorHandler
        
        nroSerie = Trim(CStr(wsActive.Cells(row, "B").Value))           ' Columna B
        orderService = Trim(CStr(wsActive.Cells(row, "D").Value))       ' Columna D
        textoDescriptivo = Trim(CStr(wsActive.Cells(row, "E").Value))   ' Columna E
        hsTrabajadas = Trim(CStr(wsActive.Cells(row, "F").Value))       ' Columna F
        repuestos = Trim(CStr(wsActive.Cells(row, "G").Value))          ' Columna G

        seguridadElectrica = Trim(CStr(wsActive.Cells(row, "H").Value)) ' Columna H
        rigel = Trim(CStr(wsActive.Cells(row, "I").Value))              ' Columna I
        pts = Trim(CStr(wsActive.Cells(row, "J").Value))                ' Columna J

        '––– Iniciar IW32 –––
        session.StartTransaction "IW32"
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        HandleSAPPopups session, r
        
        '--- Chequeo bandera verde ---
        sysStatus = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT").Text
        If InStr(1, UCase(sysStatus), "CRTD") > 0 Then
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").Text = "t02"
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
        Debug.Print fieldValue
        If (Trim(fieldValue) = "") Then
           session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-LARNT[16,0]").Text = actType
           session.findById("wnd[0]").sendVKey 0
           HandleSAPPopups session, r
        End If
        
        ' Verificar si hay "ok" en alguna de las variables antes de entrar a equipment
        If LCase(rigel) = "ok" Or LCase(seguridadElectrica) = "ok" Or LCase(pts) = "ok" Then
        Debug.Print "dentro del equipment"
        
            'entrar a equipment
            'session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0,4]").SetFocus
            'session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0,4]").caretPosition = 0
            'session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").getAbsoluteRow(4).Selected = True
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/" & _
                        "ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
        
            ' Determinar códigos de equipo
            infoCodigo = ""
            infoCodigo2 = ""
            doubleEquip = False
            Select Case True
                Case LCase(rigel) = "ok" And LCase(seguridadElectrica) = "ok"
                    doubleEquip = True: infoCodigo = codigoRigel: infoCodigo2 = codigoSE
                    Debug.Print "en 1er caso"
                Case LCase(rigel) = "ok"
                    infoCodigo = codigoRigel
                    Debug.Print "en 2do caso"
                Case LCase(pts) = "ok" And LCase(seguridadElectrica) = "ok"
                    doubleEquip = True: infoCodigo = codigoPTS: infoCodigo2 = codigoSE
                    Debug.Print "en 3er caso"
                Case LCase(pts) = "ok"
                    infoCodigo = codigoPTS
                    Debug.Print "en 4to caso"
                Case LCase(seguridadElectrica) = "ok"
                    infoCodigo = codigoSE
                    Debug.Print "en 5to caso"
            End Select
            
            Debug.Print infoCodigo
           
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
                    Debug.Print "doubleEquip"
                ElseIf infoCodigo <> "" Then
                    session.findById("wnd[1]/tbar[0]/btn[5]").press
                    session.findById("wnd[1]/usr/ctxtAFFHD-EQUNR").Text = infoCodigo
                    session.findById("wnd[1]").sendVKey 0
                    session.findById("wnd[1]/tbar[0]/btn[29]").press
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/tbar[0]/btn[3]").press
                    Debug.Print "infocodigo"
                End If
            Else
                ' Ya hay equipos cargados, no agregar
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                Debug.Print
            End If
        End If
        
        'repuestos
        If repuestos <> "" Then
            Call cargarRepuestos(session, row, repuestos)
        End If
        
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


Sub cargarRepuestos( _
    ByRef session As Object, _
    ByVal rowExcel As Long, _
    ByVal repuestos As String)

    Dim wsMP As Worksheet
    Dim material As String, qty As String, batch As String
    Dim sloc As String, planta As String
    Dim baseId As String
    Dim rowIndex As Long
    Dim partes() As String
    Dim k As Long
    Dim nroRepuesto As Long
    Dim rngBusca As Range
    Dim lastRow As Long
    Dim existingMat As String
    Dim repRow As Long
    
    ' Hoja donde está la tabla de repuestos (col L:P)
    Set wsMP = ThisWorkbook.Sheets("MP")
    
    ' Planta la tomamos de pestaña DATOS (col D)
    planta = Trim$(CStr(ThisWorkbook.Sheets("datos").Cells(3, "D").Value))
    
    ' Si no hay nada en repuestos, salimos
    repuestos = Trim$(repuestos)
    If repuestos = "" Then Exit Sub
    
    ' Abrir pestaña de componentes (MUEB)
    session.findById( _
        "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/" & _
        "ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB" _
    ).Select
    
    baseId = "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/" & _
             "ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/" & _
             "ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/"
    
    rowIndex = 0

    ' Separar "1,2,3" ? partes(0)="1", partes(1)="2", partes(2)="3"
    partes = Split(repuestos, ",")
    
    ' Última fila usada en la columna L (tabla de repuestos)
    lastRow = wsMP.Cells(wsMP.Rows.Count, "L").End(xlUp).row
    
    For k = LBound(partes) To UBound(partes)
        
        nroRepuesto = CLng(Val(Trim$(partes(k))))
        If nroRepuesto = 0 Then
            Debug.Print "Valor de repuesto inválido en lista: "; partes(k)
            GoTo SiguienteRepuesto
        End If
        
        ' Buscar nroRepuesto en columna L, desde fila 4 hasta lastRow
        Set rngBusca = wsMP.Range("L4:L" & lastRow).Find( _
            What:=nroRepuesto, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)
        
        If rngBusca Is Nothing Then
            Debug.Print "Repuesto " & nroRepuesto & " no encontrado en col L."
            GoTo SiguienteRepuesto
        End If
        
        ' Fila de la tabla de repuestos
        repRow = rngBusca.row
        
        ' Leer datos de esa fila: M (Material), N (Batch), O (Qty), P (SLoc)
        material = Trim$(CStr(wsMP.Cells(repRow, "M").Value))
        batch = Trim$(CStr(wsMP.Cells(repRow, "N").Value))
        qty = Trim$(CStr(wsMP.Cells(repRow, "O").Value))
        sloc = Trim$(CStr(wsMP.Cells(repRow, "P").Value))
        
        ' Validación mínima
        If material = "" Then
            Debug.Print "Fila " & repRow & " repuesto " & nroRepuesto & " sin material, se omite."
            GoTo SiguienteRepuesto
        End If
        
        If qty = "" Then qty = "1"
        
        ' Log de debug
        Debug.Print "SAP row=" & rowIndex & " | rep=" & nroRepuesto & _
                    " | Mat=" & material & " | Batch=" & batch & _
                    " | Qty=" & qty & " | SLoc=" & sloc & " | Planta=" & planta
        
        ' Verificar si ya hay un material en esa fila de SAP
        existingMat = Trim$( _
            session.findById(baseId & "ctxtRESBD-MATNR[1," & rowIndex & "]").Text _
        )
        Debug.Print existingMat
        
        If existingMat <> "" Then
            Debug.Print "Fila SAP " & rowIndex & " ya tiene material (" & existingMat & "), se saltea."
            rowIndex = rowIndex + 1
            GoTo SiguienteRepuesto
        End If
        
        ' Cargar nuevo componente en SAP
        With session
            .findById(baseId & "ctxtRESBD-MATNR[1," & rowIndex & "]").Text = material
            .findById(baseId & "txtRESBD-MENGE[4," & rowIndex & "]").Text = qty
            .findById(baseId & "ctxtRESBD-LGORT[8," & rowIndex & "]").Text = sloc
            .findById(baseId & "ctxtRESBD-WERKS[9," & rowIndex & "]").Text = planta
            .findById(baseId & "txtRESBD-VORNR[10," & rowIndex & "]").Text = "0010"
            .findById(baseId & "ctxtRESBD-CHARG[11," & rowIndex & "]").Text = batch
        End With
        
        Debug.Print ">>> Repuesto " & nroRepuesto & " cargado en fila SAP " & rowIndex
        rowIndex = rowIndex + 1
        
SiguienteRepuesto:
    Next k

End Sub