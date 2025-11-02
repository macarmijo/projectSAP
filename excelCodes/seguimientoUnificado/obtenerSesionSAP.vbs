Function ObtenerSesionSAP() As Object
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object

    On Error GoTo ManejarError

    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0)
    Set session = connection.Children(0)

    Set ObtenerSesionSAP = session
    Exit Function

ManejarError:
    MsgBox "No se pudo establecer la conexión con SAP. Verifique que SAP esté abierto.", vbCritical
    Set ObtenerSesionSAP = Nothing
End Function
