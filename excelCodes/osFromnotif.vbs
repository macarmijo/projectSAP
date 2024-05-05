Private Sub CommandButton1_Click()
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim r As Range, c As Range
    Dim lastRow As Long
    
    ' Set up SAP GUI connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)
    
    ' Find the last row with data in column A
    lastRow = ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).Rows.Count, "A").End(xlUp).Row
    
    ' Set the range from A2 to the last used row in column A
    Set r = ThisWorkbook.Sheets(1).Range("A2:A" & lastRow)
    
    For Each c In r
        session.StartTransaction ("IW53")
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = c.Value
        session.findById("wnd[0]").sendVKey 0
        
        c.Next(1, 1).Value = session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text
    Next c
End Sub