Private Sub GetPrinterList(ctl As Control)
    'provides list of avilable printers
    Dim prt As Printer
    For Each prt In Printers
        ctl.AddItem prt.DeviceName
    Next prt
    ctl = Application.Printer.DeviceName
End Sub
