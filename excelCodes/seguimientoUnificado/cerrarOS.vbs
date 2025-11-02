Sub cerrarOS()
    Dim session As Object
    Dim r As Range, c As Range
    Dim row As Long
    Dim orderService As String
    Dim wsActive As Worksheet
    Dim popupPresent As Boolean
    Dim maxAttempts As Integer, attempt As Integer

     ' Set up SAP GUI connection
    Set session = ObtenerSesionSAP()
    If session Is Nothing Then Exit Sub
    
    Set wsActive = ThisWorkbook.Sheets("bot-hyperautomate") ' hoja específica

    row = 3
    Do While Not UCase(Trim(wsActive.Cells(row, 2).Value)) = "END" And wsActive.Cells(row, 2).Value <> ""
        orderService = Trim(CStr(wsActive.Cells(row, 2).Value))

        On Error GoTo WarningHandler

        session.StartTransaction "IW32"
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0

        ' Ejecutar TECO
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").SetFocus
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[36]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]").sendVKey 0

        On Error GoTo 0
        GoTo ContinueLoop

WarningHandler:
        popupPresent = True
        attempt = 0
        maxAttempts = 5

        While popupPresent And attempt < maxAttempts
            attempt = attempt + 1
            MsgBox "Popup detected. Please handle it manually and then click OK to continue.", vbExclamation, "SAP Popup Warning"
            On Error Resume Next
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            If Err.Number = 0 Then popupPresent = False
            On Error GoTo WarningHandler
        Wend

        If popupPresent Then
            MsgBox "The popup couldn't be closed automatically. Please check the process.", vbCritical, "Error"
            Exit Sub
        End If
        Resume

ContinueLoop:
        On Error GoTo WarningHandler
        row = row + 1
    Loop

    session.StartTransaction "IW52"
    MsgBox "Tarea completada.", vbInformation
    Exit Sub

End Sub

'created by Maca Armijo
