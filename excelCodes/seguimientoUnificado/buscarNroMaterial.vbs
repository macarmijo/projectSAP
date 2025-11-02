Sub buscarNroMaterial()
    Dim session As Object
    Dim r As Range, c As Range
    Dim lastRow As Long
    Dim plantCov As String, plantCovZ As String, plantMdt As String, plantMdtZ As String, actType As String
    
    ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    ' Fijar hoja específica para trabajar
    Set wsActive = ThisWorkbook.Sheets("smartsheet")
    ' Find the last row with data in column G (7)
    lastRow = wsActive.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "G").End(xlUp).row
    ' Set the range from G4 to the last used row in column G
    Set r = wsActive.Range("G4:G" & lastRow)
    
    plantCov = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(3, 8).Value))
    plantMdt = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(3, 9).Value))
    actType = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(3, 10).Value))
    plantCovZ = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(3, 11).Value))
    plantMdtZ = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(3, 12).Value))

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
        
        ' Input plant number to filter for Specific Country
        session.findById("wnd[0]/usr/ctxtWERK-LOW").Text = plantCov
        session.findById("wnd[0]/usr/btn%_WERK_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = plantCovZ
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = plantMdt
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = plantMdtZ
   
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
            session.findById("wnd[0]/usr/ctxtRISA0-MATNR").Text = ""
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
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    Exit Sub
    
    ' Display a completion message
    MsgBox "Tarea completada.", vbInformation, "Información"
    ' Return to SAP home screen
    session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen

End Sub

'created by Maca Armijo
