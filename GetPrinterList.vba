Private Sub GetPrinterList(ctl As Control)
    'provides list of avilable printers
    'to combo or list box control
    Dim prt As Printer
    For Each prt In Printers
        ctl.AddItem prt.DeviceName
    Next prt
    ctl = Application.Printer.DeviceName
End Sub


;------------------------------------------------------------------------------
;Code to include in form

Private Sub Form_Load()
    Call GetPrinterList(cbxPrinterList)
End Sub

Private Sub cmdChosen_Click()
    ;The code opens the report hidden and sets the Printer property of the report
    ;to be the report that you selected. Then, the code opens the report again,
    ;this time in normal view, which causes Access to print the report.
    On Error Resume Next
    
    Dim reportName As String
    reportName = cboObjects.Value
    
    DoCmd.OpenReport reportName, View:=acPreview, WindowMode:=acHidden
    Set Reports(reportName).Printer = _
      Application.Printers(cboDestination.ListIndex)
    DoCmd.OpenReport reportName, View:=acViewNormal
End Sub
