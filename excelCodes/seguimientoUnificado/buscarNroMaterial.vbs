Sub buscarNroMaterial()
    Dim session As Object
    Dim i As Long
    Dim serialNumber As String, materialnro As String
    Dim plantCov As String, plantCovZ As String, plantMdt As String, plantMdtZ As String
    Dim wsActive As Worksheet

    ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub

    ' Fijar hoja específica para trabajar
    Set wsActive = ThisWorkbook.Sheets("smartsheet")

    plantCov = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(11, 2).Value))
    plantMdt = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(11, 3).Value))
    plantCovZ = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(11, 4).Value))
    plantMdtZ = Trim(CStr(ThisWorkbook.Sheets("datos").Cells(11, 5).Value))
    
    i = 3  ' Primera fila con datos en columna D

    ' LOOP SEGURO HASTA "END"
    Do While UCase(Trim(CStr(wsActive.Cells(i, "D").Value))) <> "END"

        ' Si la celda está vacía ? fin del loop
        If Trim(CStr(wsActive.Cells(i, "D").Value)) = "" Then Exit Do

        ' Obtener serial number
        serialNumber = Trim(CStr(wsActive.Cells(i, "D").Value))

        ' Truncar a últimos 10 caracteres
        If Len(serialNumber) > 10 Then
            serialNumber = Right(serialNumber, 10)
        End If

        ' ---------- IQ09 ----------
        session.StartTransaction ("IQ09")
        session.findById("wnd[0]/usr/txtSERNR-LOW").Text = serialNumber

        ' Filtro de plantas
        session.findById("wnd[0]/usr/ctxtWERK-LOW").Text = plantCov
        session.findById("wnd[0]/usr/btn%_WERK_%_APP_%-VALU_PUSH").press

        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" & _
                         "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = plantCovZ
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" & _
                         "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = plantMdt
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" & _
                         "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = plantMdtZ

        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/tbar[1]/btn[8]").press

        ' Intentar extraer número de material
        On Error Resume Next
        materialnro = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\07/" & _
                        "ssubSUB_DATA:SAPLITO0:0122/subSUB_0122A:SAPLITO0:1521/ctxtITOB-MATNR").Text
        On Error GoTo 0

        If materialnro <> "" Then
            wsActive.Cells(i, "C").Value = materialnro    ' escribe en columna F
        Else
            ' ---------- IQ03 fallback ----------
            session.StartTransaction ("IQ03")
            session.findById("wnd[0]/usr/ctxtRISA0-MATNR").Text = ""
            session.findById("wnd[0]/usr/ctxtRISA0-SERNR").Text = serialNumber
            session.findById("wnd[0]").sendVKey 0

            On Error Resume Next
            materialnro = session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/" & _
                            "subSUB_0152A:SAPLITO0:1521/ctxtITOB-MATNR").Text
            On Error GoTo 0

            If materialnro <> "" Then
                wsActive.Cells(i, "C").Value = materialnro
            Else
                wsActive.Cells(i, "C").Value = "Material No Encontrado"
            End If
        End If

        i = i + 1
    Loop

    MsgBox "Tarea completada.", vbInformation, "Información"
    session.findById("wnd[0]").sendVKey 15

End Sub