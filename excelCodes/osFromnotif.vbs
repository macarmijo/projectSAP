Sub buscarNroOS()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim r As Range, c As Range
    Dim lastRow As Long
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    
    ' Find the last row with data in column A
    lastRow = ThisWorkbook.ActiveSheet.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range from A4 to the last used row in column A
    Set r = ThisWorkbook.ActiveSheet.Range("A4:A" & lastRow)
    
    For Each c In r
        
        session.StartTransaction ("IW52")
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = c.Value
        session.findById("wnd[0]").sendVKey 0
    
        ' Retrieve values from SAP
        Dim os As String
        os = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text
    
        ' Check if the retrieved values are empty
        If os <> "" Then
            c.Offset(0, 1).Value = os
        Else
            c.Offset(0, 1).Value = "OS sin crear"
        End If
        
        ' Return to SAP home screen
        session.findById("wnd[0]").sendVKey 15  ' VKey 15 is typically used to go back to the home screen or initial screen

End Sub
