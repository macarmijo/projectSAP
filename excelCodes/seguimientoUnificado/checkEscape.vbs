Option Explicit

' Manejo universal para detectar la tecla ESC con opcion, continuo o freno script.
Public Function CheckEscape(ByVal errNumber As Long) As Boolean
    If errNumber = 18 Then
        Dim resp As VbMsgBoxResult

        resp = MsgBox( _
            "Se detectó la tecla ESC." & vbCrLf & vbCrLf & _
            "¿Desea detener el script?", _
            vbYesNo + vbQuestion, _
            "STOP SCRIPT")

        If resp = vbYes Then
            MsgBox "El script ha sido detenido con exito.", vbInformation, "Proceso cancelado"
            CheckEscape = True     ' Cortar el macro
        Else
            CheckEscape = False    ' Continuar el macro
        End If

    Else
        CheckEscape = False
    End If
End Function

