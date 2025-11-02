'created by Maca Armijo
Sub IW42()
    Dim session As Object
    Dim r As Range, c As Range
    Dim orderService As Long
    Dim numeroPersona As String
    Dim pestanaExcel As Worksheet
    Dim i As Long
    Dim cellValue As String, actWorkValue As String
    Dim popupType As String
         
    ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    MsgBox "Presione OK si desea consumir las horas de las OS correspondientes (IW42)"
    
    ' Set the active worksheet to pestanaExcel
    Set pestanaExcel = ThisWorkbook.Sheets("bot-hyperautomate")
    numeroPersona = Trim(CStr(pestanaExcel.Cells(3, 10).Value))
    
    i = 3 ' Start from row 2
    Do While True
        cellValue = Trim(CStr(pestanaExcel.Cells(i, 2).Value))
        
        ' Check if cellValue is empty, "END", "end", or a non-numeric value
        If cellValue = "" Or LCase(cellValue) = "end" Or Not IsNumeric(cellValue) Then
            Exit Do
        End If
        
        orderService = Trim(CStr(pestanaExcel.Cells(i, 2).Value))
        
        On Error GoTo GeneralErrorHandler
        
        'consumo horas trabajadas
        session.StartTransaction ("IW42")
        session.findById("wnd[0]/usr/subHEADER:SAPLCMFU:0201/ctxtCMFUD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        actWorkValue = session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/txtAFRUD-ISMNW_2[4,0]").Text
        
        ' Verificar si actWorkValue está vacío
        If actWorkValue <> "" Then
            ' Asignar el número de persona al campo correspondiente en SAP
            session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").Text = numeroPersona
            session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").SetFocus
            session.findById("wnd[0]/usr/subSUB1:SAPLCMFU:0011/subSUB11_1:SAPLCMFU:0101/subOBJECTSCREEN_01:SAPLCORU:3360/tblSAPLCORUTABCNTR_3360/ctxtAFRUD-PERNR[14,0]").caretPosition = 7
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            
            ' Handle SAP popups
            HandleSAPPopups session
        End If
        
        ' Continue to the next row
        On Error GoTo 0
        GoTo ContinueLoop
    
GeneralErrorHandler:
        ' Log the error in Excel
        pestanaExcel.Cells(i, 3).Value = "Error: " & Err.Description
        Resume ContinueLoop
    
ContinueLoop:
        i = i + 1
    Loop
    
    ' Return to SAP home screen
    session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen
    MsgBox "Horas consumidas con éxito."
    Exit Sub

End Sub

Sub HandleSAPPopups(ByRef session As Object)
    ' Check for the specific popup and press Enter to close it
    If session.Children.Count > 1 Then
        popupType = session.findById("wnd[1]/usr/txtMESSTXT1").Text
        If popupType Like "*Customer*block*" Then
            session.findById("wnd[1]/tbar[0]/btn[0]").press ' Press Enter to close popup
        End If
    End If
End Sub
