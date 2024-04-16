'Este busca jugar con las distintas opciones de conexion con excel

 'Conexion con hoja de excel activa de donde saco y guardo informacion
 Dim objExcel
 Dim objSheet, intRow, i
 Set objExcel = GetObject(,"Excel.Application")
 Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
 Set book1 = objExcel.Worksheets(1)
 Set book2 = objExcel.Worksheets(2)
 Set book3 = objExcel.Worksheets(3)


 'defino variables que van a obtener sus valores de la hoja de excel activa al momento de correr el script
 'fila , col
 WScript.Echo "probando"

 equipoSeguridadElectrica = "1000001149900"
 'directorio podriamos sacarlo del excel tambien!! VER de modificar!
 directorio = "C:\Users\Roldam13\OneDrive - Medtronic PLC\Servicio Tecnico\Comercial\Contratos de Mantenimiento\Hospital Italiano\ABRIL 2024"

 'ESTO FUNCIONA! boook1 es la pestaña Sheet1 de excel- book2 es la pestaña Sheet2 etc!
book1.Cells(15, 1).Value = "hola"
book2.Cells(1, 1).Value = "funcionooooo"