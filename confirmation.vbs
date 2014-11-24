REM Check for installed printer and output confirmation message.
    REM MsgBox printerExists("\\rno-g-wds01\RNO-G-Frenzy-100")
    printerExists("\\rno-g-wds01\RNO-G-Frenzy-100")

    Function printerExists(str)
        printerExists = False 
        Dim objWMIService
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

        Dim colPrinters
        Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

        Dim objPrinter
        For Each objPrinter In colPrinters
            If objPrinter.Name = str Then
                printerExists = True
                Exit For
            End If
        Next

        If printerExists = false Then
            MsgBox "Error!"
        else
            MsgBox "Success!"
        End If

    End Function

    