Sub getOSfromNotif()
    'defino variables
    Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
    Dim Sheet1 As Worksheet
    Dim valorCelda As String
    Dim i As Long
    Dim notif As String, os As String
    
    'config integracion SAP - usar SAP abierto!
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapGuiApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SapGuiApp.Children(0)
    Set session = Connection.Children(0)

    Set Sheet1 = ThisWorkbook.Sheets(1)
    
    For i = 2 To Sheet1.UsedRange.Rows.Count
        notif = Trim(CStr(Sheet1.Cells(i, 1).Value))
       Debug.Print notif
       ' Start the IW52 transaction
        session.StartTransaction ("IW52")
        
        ' Enter the notification number in SAP
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = notif
        session.findById("wnd[0]").sendVKey 0
        
        ' Get the OS number from SAP and save it in Excel
        os = Trim(CStr(session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-AUFNR").Text))
        Sheet1.Cells(i, 2).Value = os
        
        ' Move to the next row in Excel
        session.findById("wnd[0]").sendVKey 0
    Next i
    
End Sub
