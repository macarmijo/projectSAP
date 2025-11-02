Sub DescargarMultiplePDF()
    Dim session As Object
    Dim ws As Worksheet
    Dim i As Long
    Dim orden As String
    
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("verificacion")
    
    i = 3 ' Fila inicial
    
    Do While UCase(Trim(CStr(ws.Cells(i, 3).Value))) <> "END"
        orden = Trim(CStr(ws.Cells(i, 3).Value))
        If orden <> "" Then
            Debug.Print "Descargando OS " & orden
            
            ' Abrir NZCSOUTPUT
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzcsoutput"
            session.findById("wnd[0]").sendVKey 0
            
            ' Completar número de orden
            session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").Text = orden
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            
            ' Presionar relojito (Imprimir)
            session.findById("wnd[1]/usr/btnBUTTON_2").press
            
            ' Elegir impresora local
            session.findById("wnd[1]/usr/ctxtSSFPP-TDDEST").Text = "locl"
            session.findById("wnd[1]/tbar[0]/btn[8]").press
            
            ' Comando PDF!
            session.findById("wnd[0]/tbar[0]/okcd").Text = "pdf!"
            session.findById("wnd[0]").sendVKey 0
            
            ' ?? Pausa manual — Espera a que el usuario descargue el archivo
            MsgBox "Guardá el archivo PDF manualmente y hacé clic en Aceptar para continuar con la próxima OS.", vbInformation, "Descargar PDF"
            
            ' Volver a NZCSOUTPUT
            'session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzcsoutput"
            'session.findById("wnd[0]").sendVKey 0
            
        End If
        i = i + 1
    Loop
    
    MsgBox "Descarga múltiple finalizada.", vbInformation
End Sub