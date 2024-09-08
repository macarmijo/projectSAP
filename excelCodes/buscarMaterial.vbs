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
        
        ' Ensure the value is at most 10 characters
        Dim serialNumber As String
        serialNumber = c.Value
        If Len(serialNumber) > 10 Then
            serialNumber = Right(serialNumber, 10)
        End If
    
        ' Start the IQ09 transaction
        session.StartTransaction ("IQ09")
        ' Input serial number
        session.findById("wnd[0]/usr/txtSERNR-LOW").Text = serialNumber
        
        ' Input plant number to filter for Argentina
        session.findById("wnd[0]/usr/ctxtWERK-LOW").Text = "1394"
        session.findById("wnd[0]/usr/btn%_WERK_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "z394"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "1164"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "z164"
   
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/tbar[1]/btn[8]").press

        ' Retrieve values from SAP
        Dim materialnro As String
        materialnro = ""
        
        On Error Resume Next
        materialnro = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\07/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122A:SAPLITO0:1521/ctxtITOB-MATNR").Text
        On Error GoTo 0

        ' Check if the retrieved value is empty
        If materialnro <> "" Then
            c.Offset(0, -1).Value = materialnro
        Else
            ' Start another transaction (IQ02) if material number is not found in IQ09
            session.StartTransaction ("IQ03")
            session.findById("wnd[0]/usr/ctxtRISA0-SERNR").Text = serialNumber
            session.findById("wnd[0]").sendVKey 0
            
            On Error Resume Next
            materialnro = session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152A:SAPLITO0:1521/ctxtITOB-MATNR").Text
            On Error GoTo 0
            
            If materialnro <> "" Then
                c.Offset(0, -1).Value = materialnro
            Else
                c.Offset(0, -1).Value = "Material No Encontrado"
            End If
        End If

    Next c
    
    ' Display a completion message
    MsgBox "Tarea completada.", vbInformation, "Información"
    ' Return to SAP home screen
    session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen
    Exit Sub

ErrorHandler:
    MsgBox "Necesitas abrir SAP.", vbCritical, "Error de Conexión"
End Sub