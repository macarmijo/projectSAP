Sub consumirHoras()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim r As Range, c As Range
    Dim orderService As Long
    Dim numeroPersona As String
    Dim pestanaExcel As Worksheet
    Dim i As Long
    Dim cellValue As String, actWorkValue As String
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    MsgBox "Presione OK si desea consumir las horas de las OS correspondientes (IW42)"
    
    ' Set the active worksheet to pestanaExcel
    Set pestanaExcel = ThisWorkbook.ActiveSheet
    numeroPersona = Trim(CStr(pestanaExcel.Cells(2, 17).Value))
    'pestanaExcel.Cells(2, 12).Value = numeroPersona
    
    
    i = 2 ' Start from row 2
    Do While True
        cellValue = Trim(CStr(pestanaExcel.Cells(i, 1).Value))
        
        ' Check if cellValue is empty, "END", "end", or a non-numeric value
        If cellValue = "" Or LCase(cellValue) = "end" Or Not IsNumeric(cellValue) Then
            Exit Do
        End If
        
        orderService = Trim(CStr(pestanaExcel.Cells(i, 2).Value))
        
        'consumo horas trabajadas
        session.StartTransaction ("IW42")
        session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0201/ctxtCMFUD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        actWorkValue = session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/txtAFRUD-ISMNW_2[4,0]").Text
        
        
        ' Verificar si actWorkValue está vacío
        If actWorkValue <> "" Then
            'pestanaExcel.Cells(2, 12).Value = actWorkValue
            ' MsgBox "No se encontró 'act. work' para la orden de servicio: " & orderService
            ' Esto quedo en caso que querramos ver que orden no tiene horas cargadas
            ' Asignar el número de persona al campo correspondiente en SAP
            session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").Text = numeroPersona
            session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").SetFocus
            session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").caretPosition = 7
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[11]").press
        End If
        
        i = i + 1
    Loop
    
    ' Return to SAP home screen
    session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen
    MsgBox "Horas consumidas con éxito."
    
End Sub
