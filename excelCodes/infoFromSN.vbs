Private Sub CommandButton1_Click()
    'defino variables
        Dim SapGuiAuto As Object, SapGuiApp As Object, Connection As Object, session As Object
        Dim Sheet3 As Worksheet
        Dim valorCelda As String
        Dim i As Long
        Dim notif As String, os As String
        
        'config integracion SAP - usar SAP abierto!
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SapGuiApp = SapGuiAuto.GetScriptingEngine
        Set Connection = SapGuiApp.Children(0)
        Set session = Connection.Children(0)
    
        Set Sheet3 = ThisWorkbook.Sheets(3)
        
        For i = 2 To Sheet3.UsedRange.Rows.Count
            nroSerie = Trim(CStr(Sheet3.Cells(i, 1).Value))
           
           ' Start the IW52 transaction
            session.StartTransaction ("IQ02")
            
            ' Enter the serial number & no material in SAP
            session.findById("wnd[0]/usr/ctxtRISA0-MATNR").Text = ""
            session.findById("wnd[0]/usr/ctxtRISA0-SERNR").Text = nroSerie
            session.findById("wnd[0]").sendVKey 0
            
            Sheet3.Cells(i, 2).Value = session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152A:SAPLITO0:1521/ctxtITOB-MATNR").Text
            Sheet3.Cells(i, 3).Value = session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152B:SAPLITO0:1525/txtITOB-SHTXT").Text
            Sheet3.Cells(i, 4).Value = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\07/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122D:SAPLITO0:1222/subSUB_1222A:SAPLIPAR:0801/txtDIADR-NAME_LIST").Text
            Sheet3.Cells(i, 5).Value = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\07/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122D:SAPLITO0:1222/subSUB_1222A:SAPLIPAR:0801/ctxtIHPA-PARNR").Text
        Next i
    End Sub
    