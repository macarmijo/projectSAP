Sub buscarOS()
    Dim session As Object
    Dim wsActive As Worksheet
    Dim r As Range, c As Range
    Dim lastRow As Long
    
    ' Fijar hoja específica para trabajar
    Set wsActive = ThisWorkbook.Sheets("smartsheet")
    
    ' Find the last row with data in column A
    lastRow = ThisWorkbook.ActiveSheet.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "A").End(xlUp).row
    
    ' Set the range from A2 to the last used row in column A
    Set r = ThisWorkbook.ActiveSheet.Range("A4:A" & lastRow)
    
    ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
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

'created by Maca Armijo
