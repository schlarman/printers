' Add "Frenzy" Printer - Desktop Support Reno - CustomInk 2014
' http://thydzik.com/vbavbs-check-if-printer-is-installed/

' Popup message confirming installation.
intResponse = Msgbox("Would you like to add the Frenzy printer?", vbYesNo, "Confirm Add Printer")

' Set path for network printer "Frenzy".
uncpath = "\\rno-g-wds01\RNO-G-Frenzy-100"

' If "Yes" add printer.
If intResponse = vbYes Then
  Set objNet = CreateObject("WScript.Network")
  objNet.AddWindowsPrinterConnection uncpath

    ' Check for installed printer and output confirmation message.
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

        ' Show confirmation message.
        If printerExists = false Then
            MsgBox "Oops, the printer did not install. Please contact the Desktop Support."
        else
            MsgBox "Success! You can close this window."
        End If

    End Function

' If user selects "No"
Else
    Msgbox "No changes were made."
End If