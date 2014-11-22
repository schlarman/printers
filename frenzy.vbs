REM Add "Frenzy" Printer - Torin Emard - CustomInk 2014
REM http://thydzik.com/vbavbs-check-if-printer-is-installed/

intResponse = Msgbox("Would you like to add the Frenzy printer?", vbYesNo, "Confirm Add Printer")

uncpath = "\\rno-g-wds01\RNO-G-Frenzy-100"

If intResponse = vbYes Then
  Set objNet = CreateObject("WScript.Network")
  objNet.AddWindowsPrinterConnection uncpath

  REM Check for installed printer and output confirmation message.

MsgBox printerExists("\\rno-g-wds01\RNO-G-Frenzy-100")

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

End Function

Else
    Msgbox "No changes were made."
End If

