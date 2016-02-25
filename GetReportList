Private Sub GetReportList(ctl As Control)
    'Provides a list of available reports in your application..
    Dim item As AccessObject
    
    ' Clear the list before adding new items.
    ctl.RowSourceType = "Value List"
    ctl.RowSource = vbNullString
    For Each item In CurrentProject.AllReports
        ctl.AddItem item.Name
    Next item
End Sub...
