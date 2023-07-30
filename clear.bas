Sub clearAllDataFormatsAndDeleteSheetsExceptActiveSheet()
    Application.DisplayAlerts = False
    Dim activeSheetName As String
    activeSheetName = ActiveSheet.Name

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> activeSheetName Then
            Application.DisplayAlerts = False ' Avoid confirmation prompt
            ws.Delete
            Application.DisplayAlerts = True
        Else
            On Error Resume Next
            ws.Cells.ClearContents
            ws.Cells.ClearFormats
            On Error GoTo 0
        End If
    Next ws

    Application.DisplayAlerts = True
    MsgBox "All data, formats, and sheets have been cleared except the active sheet.", vbInformation
End Sub
