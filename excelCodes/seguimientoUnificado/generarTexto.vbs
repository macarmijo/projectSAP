Function GenerarTexto(bateria As String, o2 As String, atg As String, se As String, kit As String, horas As Variant, Optional extraTexto As String = "") As String
    Dim partes As String
    Dim textoBase As String
    Dim b As String, m As String, a As String, s As String, k As String

    ' Convertimos todo a minúsculas para comparación
    b = LCase(bateria)
    m = LCase(o2)
    a = LCase(atg)
    s = LCase(se)
    k = LCase(kit)

    ' CASOS ESPECIALES: todos iguales
    If b = "x" And m = "x" And a = "x" And s = "x" And k = "x" Then
        GenerarTexto = "Se retira equipo - No pasa pruebas de mantenimiento preventivo"
        Exit Function
    End If

    If b = "mdt" And m = "mdt" And a = "mdt" And s = "mdt" And k = "mdt" Then
        GenerarTexto = "Equipo en taller Medtronic"
        Exit Function
    End If

    If b = "na" And m = "na" And a = "na" And s = "na" And k = "na" Then
        GenerarTexto = "Equipo en uso, no se pudo realizar mantenimiento preventivo programado. Reagendar visita."
        Exit Function
    End If

    If b = "nn" And m = "nn" And a = "nn" And s = "nn" And k = "nn" Then
        GenerarTexto = "Equipo no encontrado"
        Exit Function
    End If

    ' CASO ESPECIAL: no, no, ok, ok, no
    If b = "no" And m = "no" And a = "ok" And s = "ok" And k = "no" Then
        textoBase = "Equipo con " & horas & " horas. Se realizó ATG y test de seguridad eléctrica. Equipo operativo."
        If Trim(extraTexto) <> "" Then textoBase = textoBase & " " & extraTexto
        GenerarTexto = textoBase
        Exit Function
    End If

    ' CASO GENERAL: se realizó ATG
    If a = "ok" Then
        textoBase = "Equipo con " & horas & " horas. Se realizó ATG"

        ' Detectar componentes cambiados
        If b = "ok" Then partes = partes & ", batería"
        If m = "ok" Then partes = partes & ", celda O2"
        If k = "ok" Then partes = partes & ", kit preventivo 10k horas"
        If s = "ok" Then partes = partes & ", test de seguridad eléctrica"

        If partes <> "" Then
            textoBase = textoBase & ", cambio de partes" & partes
        End If

        textoBase = textoBase & ". Equipo operativo."
    Else
        textoBase = "Equipo con " & horas & " horas. Sin mantenimiento registrado."
    End If

    ' Agregar texto adicional si existe
    If Trim(extraTexto) <> "" Then
        textoBase = textoBase & " " & extraTexto
    End If

    GenerarTexto = textoBase
End Function