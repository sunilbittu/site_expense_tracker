Attribute VB_Name = "ExportDashboardToPDF"
Sub ExportDashboardToPDF()
    Dim ws As Worksheet
    Dim exportPath As String
    Dim fileName As String

    Set ws = ThisWorkbook.Sheets("Dashboard")

    fileName = "Dashboard_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".pdf"
    exportPath = ThisWorkbook.Path & Application.PathSeparator & fileName

    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Dashboard exported to:" & vbCrLf & exportPath, vbInformation
End Sub
