Sub buscarOS()
    On Error GoTo ErrorHandler
    application.EnableCancelKey = xlErrorHandler   ' Permite que ESC genere error 18

    Dim session As Object
    Dim wsActive As Worksheet
    Dim nroNotif As String
    Dim os As String
    Dim i As Long

    ' Hoja específica
    Set wsActive = ThisWorkbook.Sheets("MP")

    ' Conexión SAP
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub

    i = 3   ' Primera fila con datos en columna C

    ' LOOP SEGURO HASTA "END"
    Do
        nroNotif = Trim$(CStr(wsActive.Cells(i, "C").Value))

        ' Si está vacío, cortamos
        If nroNotif = "" Then Exit Do
        ' Si dice END, cortamos
        If UCase$(nroNotif) = "END" Then Exit Do

        ' ---------- IW53 ----------
        session.StartTransaction "IW53"
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = nroNotif
        session.findById("wnd[0]").sendVKey 0

        ' Leer OS
        On Error Resume Next
        os = session.findById( _
               "wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/" & _
               "txtVIQMEL-AUFNR").Text
        On Error GoTo ErrorHandler

        If os <> "" Then
            wsActive.Cells(i, "D").Value = os
        Else
            wsActive.Cells(i, "D").Value = "OS sin crear"
        End If

        i = i + 1
    Loop

    ' Volver a la pantalla inicial de SAP
    On Error Resume Next
    session.findById("wnd[0]").sendVKey 15
    On Error GoTo 0

    MsgBox "Tarea completada.", vbInformation, "Información"
    Exit Sub

' =======================
'   MANEJO DE ERRORES
' =======================
ErrorHandler:
    If CheckEscape(Err.Number) Then
        ' Cancelado por ESC
        On Error Resume Next
        session.findById("wnd[0]").sendVKey 15
        On Error GoTo 0
        Exit Sub
    Else
        If Err.Number = 18 Then
            Err.Clear
            Resume Next
        Else
            MsgBox "Error inesperado: " & Err.Description, _
                   vbCritical, "Error en buscarOS"
            Resume Next
        End If
    End If
End Sub


