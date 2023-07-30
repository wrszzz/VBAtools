Sub mergeDataWithFileSelectorAndExclusion()
    Application.ScreenUpdating = False
    Dim totalwb As Workbook, wbx As Workbook
    Dim ws As Worksheet, combinedSheet As Worksheet
    Dim path As String, filenames As String
    Dim sc As Long, i As Long, x As Long
    Dim RC As Range
    
    ' Create or get reference to the "combined" sheet
    On Error Resume Next
    Set combinedSheet = ThisWorkbook.Sheets("combined")
    On Error GoTo 0
    
    If combinedSheet Is Nothing Then
        Set combinedSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1)) ' Add as the first sheet
        combinedSheet.Name = "combined"
    End If
    
    ' Show file dialog to select files for combining
    Dim fileDialog As FileDialog
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' List of sheet names that should not be combined
    Dim excludeSheets As Variant
    excludeSheets = Array("Sheet1", "Sheet2", "Sheet3") ' Add the names of the sheets you want to exclude here
    
    With fileDialog
        .AllowMultiSelect = True
        .Title = "Select files to combine"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx", 1
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then ' User canceled the file dialog
            MsgBox "No files selected for combining.", vbExclamation
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        For Each filename In .SelectedItems
            Set wbx = Workbooks.Open(filename, UpdateLinks:=0)
            
            For i = 1 To wbx.Sheets.Count
                Set ws = wbx.Sheets(i)
                If ws.Visible = True And Not IsInArray(ws.Name, excludeSheets) Then
                    If ws.AutoFilterMode = True Then ws.AutoFilterMode = False
                    ws.Rows.RowHeight = 15
                    
                    ' Copy data
                    Dim lastRow As Long
                    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                    If lastRow > 0 Then ' Avoid copying if there's no data in the sheet
                        Dim dataRange As Range
                        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 99))
                        dataRange.Copy
                        combinedSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False ' Clear clipboard
                    End If
                End If
            Next
            
            wbx.Close False
            Set wbx = Nothing
        Next filename
    End With

    Application.ScreenUpdating = True
    MsgBox "Data merging completed!", vbInformation
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(stringToBeFound, element, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
