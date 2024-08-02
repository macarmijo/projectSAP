Sub buscarNroMaterial()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim r As Range, c As Range
    Dim lastRow As Long

    On Error GoTo ErrorHandler
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    ' If connection is established, continue with the rest of the code
    On Error GoTo 0

    ' Find the last row with data in column G (7)
    lastRow = ThisWorkbook.ActiveSheet.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "G").End(xlUp).Row

    ' Set the range from G4 to the last used row in column G
    Set r = ThisWorkbook.ActiveSheet.Range("G4:G" & lastRow)

    For Each c In r
        session.StartTransaction ("IQ09")
        'ingreso nro de serie
        session.findById("wnd[0]/usr/txtSERNR-LOW").Text = c.Value
        'session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
        
        'ingreso plant nro para filtrar las de argentina
        session.findById("wnd[0]/usr/ctxtWERK-LOW").Text = "1394"
        session.findById("wnd[0]/usr/btn%_WERK_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "z394"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "1164"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "z164"
   
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/tbar[1]/btn[8]").press

        ' Check if the desired element is present
        On Error Resume Next
        ' Retrieve values from SAP
        Dim materialnro As String
        materialnro = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\07/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122A:SAPLITO0:1521/ctxtITOB-MATNR").Text
        On Error GoTo 0

        ' Check if the retrieved values are empty
        If materialnro <> "" Then
            c.Offset(0, -1).Value = materialnro
        Else
            
            ' Example: Start another transaction (replace with your desired transaction)
            session.StartTransaction ("IQ02")
            session.findById("wnd[0]/usr/ctxtRISA0-SERNR").Text = c.Value
            session.findById("wnd[0]").sendVKey 0
            
            materialnro = session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152A:SAPLITO0:1521/ctxtITOB-MATNR").Text
            c.Offset(0, -1).Value = materialnro
        End If

        ' Return to SAP home screen
        session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen
    Next c

    ' Display a completion message
    MsgBox "Tarea completada.", vbInformation, "Información"
    
    Exit Sub

ErrorHandler:
    MsgBox "Necesitas abrir SAP.", vbCritical, "Error de Conexión"
End Sub