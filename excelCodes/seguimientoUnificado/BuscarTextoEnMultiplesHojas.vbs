Sub BuscarTextoEnMultiplesHojas()
    Dim wsOrigen As Worksheet
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim fila As Long
    Dim termino As String, filaTxt as String, columnaTxt as string
    Dim celdaEncontrada As Range
    Dim encontrado As Boolean
    
    ' Hojas
    Set wsOrigen = ThisWorkbook.Sheets("MP")
    Set ws1 = ThisWorkbook.Sheets("respis")
    Set ws2 = ThisWorkbook.Sheets("electros")
    
    fila = 3  ' fila inicial donde empiezan los términos en MP
    columnaTxt = "S"
    
    Do While UCase(Trim(CStr(wsOrigen.Cells(fila, "B").Value))) <> "END"
        
        termino = Trim(CStr(wsOrigen.Cells(fila, "B").Value))
        
        ' Si está vacío, cortamos
        If termino = "" Then Exit Do
        
        encontrado = False
        
        ' ---------- 1) Buscar en hoja "respis" ----------
        Set celdaEncontrada = Nothing
        Set celdaEncontrada = ws1.Columns("B").Find(What:=termino, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole, _
                                                    MatchCase:=False)
     
        If Not celdaEncontrada Is Nothing Then
            wsOrigen.Cells(fila, "E").Value = ws1.cells(celdaEncontrada.row, "s")
           ' wsOrigen.Cells(fila, "E").Value = "Encontrado en 'respis'!B" & celdaEncontrada.row
            encontrado = True
        Else
            ' ---------- 2) Si no está, buscar en hoja "electros" ----------
            Set celdaEncontrada = ws2.Columns("B").Find(What:=termino, _
                                                        LookIn:=xlValues, _
                                                        LookAt:=xlWhole, _
                                                        MatchCase:=False)
            If Not celdaEncontrada Is Nothing Then
                wsOrigen.Cells(fila, "E").Value = ws2.cells(celdaEncontrada.row, "s")
                'wsOrigen.Cells(fila, "E").Value = "Encontrado en 'electros'!B" & celdaEncontrada.row
                encontrado = True
            End If
        End If
        
        ' ---------- 3) Si no se encontró en ningún lado ----------
        If Not encontrado Then
            wsOrigen.Cells(fila, "E").Value = "Texto no encontrado"
        End If
        
        fila = fila + 1
    Loop
    
    MsgBox "Búsqueda completada!", vbInformation
End Sub