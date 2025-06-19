'busco a partir de la notif IW52 - 
Sub buscarOS()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim r As Range, c As Range
    Dim lastRow As Long
     
    'If connection was not possible
    On Error GoTo ErrorHandler
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    ' If connection is established, continue with the rest of the code
    On Error GoTo 0
    
    ' Find the last row with data in column A
    lastRow = ThisWorkbook.ActiveSheet.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "A").End(xlUp).row
    
    ' Set the range from A2 to the last used row in column A
    Set r = ThisWorkbook.ActiveSheet.Range("A4:A" & lastRow)
    
    For Each c In r
        session.StartTransaction ("IW53")
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = c.Value
        session.findById("wnd[0]").sendVKey 0
    
        ' Retrieve values from SAP
        Dim os As String
        os = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text
    
        ' Check if the retrieved values are empty
        If os <> "" Then
            c.Next(1, 1).Value = os
        Else
            c.Next(1, 1).Value = "OS sin crear"
        End If
    
    Next c
    
    ' Return to SAP home screen
    session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen
    ' Display a completion message
    MsgBox "Tarea completada.", vbInformation, "Información"
    
    Exit Sub

    
ErrorHandler:
    MsgBox "Necesitas abrir SAP.", vbCritical, "Error de Conexión"

End Sub