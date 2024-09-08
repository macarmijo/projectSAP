Sub cerrarOS()
    Dim r As Range, c As Range
    Dim lastRow As Long, orderService As String
    Dim pestanaExcel As Worksheet
    Dim popupPresent As Boolean
    Dim maxAttempts As Integer, attempt As Integer
    
    'If connection was not possible
    On Error GoTo ErrorHandler
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    ' If connection is established, continue with the rest of the code
    On Error GoTo 0
    
    ' Find the last row with data in column A
    lastRow = ThisWorkbook.ActiveSheet.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range from A2 to the last used row in column A
    Set r = ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow)
    ' Set the active worksheet to pestanaExcel
    Set pestanaExcel = ThisWorkbook.ActiveSheet
    
    For i = 2 To lastRow
        On Error GoTo WarningHandler
        
        orderService = Trim(CStr(pestanaExcel.Cells(i, 2).Value))
        ' Start SAP transaction IW32
        session.StartTransaction ("IW32")
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderService
        session.findById("wnd[0]").sendVKey 0
        ' Execute TECO (Technical Completion)
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").SetFocus
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-ASTTX").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[36]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]").sendVKey 0
    
        ' Continue to the next row
        On Error GoTo 0
        GoTo ContinueLoop
    
WarningHandler:
        ' Handle the popup warning
        popupPresent = True
        attempt = 0
        maxAttempts = 5 ' Number of attempts before stopping
        
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
        Resume  ' Resume the code after the user handles the popup
    
ContinueLoop:
    Next i
    
    ' Display a completion message
    MsgBox "Tarea completada.", vbInformation, "Información"
    ' Return to SAP home screen
    session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen
    Exit Sub

ErrorHandler:
    MsgBox "Necesitas abrir SAP.", vbCritical, "Error de Conexión"
    
End Sub