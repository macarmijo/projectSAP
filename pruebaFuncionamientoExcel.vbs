If Not IsObject(application) Then
    Set SapGuiAuto  = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
 End If
 If Not IsObject(connection) Then
    Set connection = application.Children(0)
 End If
 If Not IsObject(session) Then
    Set session    = connection.Children(0)
 End If
 If IsObject(WScript) Then
    WScript.ConnectObject session,     "on"
    WScript.ConnectObject application, "on"
 End If

'Este busca jugar con las distintas opciones de conexion con excel
'defino variables que van a obtener sus valores de la hoja de excel activa al momento de correr el script
'fila , col
'WScript.Echo "probando"

'ESTO FUNCIONA! boook1 es la pestaña Sheet1 de excel- book2 es la pestaña Sheet2 etc!
'book1.Cells(15, 1).Value = "hola"
'book2.Cells(1, 1).Value = "funcionooooo"

'Conexion con hoja de excel activa de donde saco y guardo informacion
Dim objExcel
Dim objSheet, intRow, i
Set objExcel = GetObject(,"Excel.Application")
Set objWorkbook = objExcel.ActiveWorkbook
Set sheet1 = objWorkbook.Worksheets(1)
Set sheet2 = objWorkbook.Worksheets(2)


   ' Función para buscar una palabra en un texto
Function buscarPalabra(texto, palabra1, palabra2, palabra3)
   If InStr(texto, palabra1) > 0 Or InStr(texto, palabra2) > 0 Or InStr(texto, palabra3) > 0 Then
      buscarPalabra = True
   Else
      buscarPalabra = False
   End If
End Function

'arranca en fila 2 - fila 1 tiene nombre de los valores de la columna
For i = 2 to sheet2.UsedRange.Rows.Count
   notif = Trim(CStr(sheet2.Cells(i, 1).Value))

   session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = notif
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]").sendVKey 0
   'guarda en excel el nro de OS
   sheet2.Cells(i, 2).Value = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text
   session.findById("wnd[0]").sendVKey 0
   
   'entro a la OS segun el nro que guarde en el paso anterior
   session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = Trim(CStr(sheet2.Cells(i, 2).Value))
   session.findById("wnd[0]").sendVKey 0
   
   'busca modelo del equipo y guardo en excel !
   sheet2.Cells(i, 3).Value = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subSUB_SERVICE:SAPLCOI3:0300/ctxtPMSDO-MATNR").text
   'guardo el nro serie del equipo
   sheet2.Cells(i, 4).Value = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subOBJECT:SAPLCOIH:7170/ctxtCAUFVD-SERIALNR").text

   ' Buscar las palabras en el texto utilizando la función buscarPalabra
   If buscarPalabra(sheet2.Cells(i, 3), "ARGON", "HT70", "5100C") Then
      sheet2.Cells(i, 5).Value = "SI"
      sheet2.Cells(i, 6).Value = "NO"
   Else
      sheet2.Cells(i, 5).Value = "SI"
      sheet2.Cells(i, 6).Value = "SI"
   End If
next
