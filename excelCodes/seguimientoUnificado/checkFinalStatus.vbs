Sub checkFinalStatus()
    Dim session As Object
    Dim wsActive As Worksheet
    Dim i As Long
    Dim notifNum As String, osNum As String
    Dim notifStat As String, osStat As String
    
    ' Conexión SAP
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    Set wsActive = ThisWorkbook.Sheets("verificacion")
    
    application.EnableCancelKey = xlErrorHandler
    On Error GoTo HandlerFatal
    
    i = 3 ' Fila inicial
    
    Do While UCase(Trim(CStr(wsActive.Cells(i, 1).Value))) <> "END"
        ' Permitir cancelar con ESC
        If Err.Number = 18 Then
            MsgBox "Proceso cancelado por el usuario (ESC).", vbExclamation
            Exit Do
        End If
        
        DoEvents
        
        ' Si la fila está vacía, salimos
        If Trim(CStr(wsActive.Cells(i, 1).Value)) = "" Then Exit Do
        
        ' --- IW53: Estado de la Notificación ---
        notifNum = Trim(CStr(wsActive.Cells(i, 1).Value)) ' Columna A
        If notifNum <> "" Then
            session.StartTransaction "IW53"
            session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = notifNum
            session.findById("wnd[0]").sendVKey 0
            notifStat = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/txtRIWO00-STTXT").Text
            wsActive.Cells(i, 2).Value = notifStat ' Columna B
        End If
        
        ' --- IW33: Estado de la Orden de Servicio ---
        osNum = Trim(CStr(wsActive.Cells(i, 3).Value)) ' Columna C
        If osNum <> "" Then
            session.StartTransaction "IW33"
            session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = osNum
            session.findById("wnd[0]").sendVKey 0
            osStat = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/" & _
                     "ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").Text
            wsActive.Cells(i, 4).Value = osStat ' Columna D
        End If
        
        i = i + 1
    Loop
    
    ' Volver a la pantalla inicial de SAP
    On Error Resume Next
    session.findById("wnd[0]").sendVKey 15
    On Error GoTo 0
    
    MsgBox "Verificación completada con éxito.", vbInformation, "Información"
    Exit Sub

' Manejo de error global
HandlerFatal:
    MsgBox "Error: " & Err.Description, vbCritical, "checkFinalStatus"
End Sub


