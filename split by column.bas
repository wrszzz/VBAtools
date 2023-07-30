Attribute VB_Name = "Ä£¿é3"
Sub SplitWorkbookBasedOnColumn()
    Dim OriginalWs As Worksheet
    Dim NewWorkbook As Workbook
    Dim NewWs As Worksheet
    Dim CurrentCell As Range
    Dim ColumnToSplit As Range
    Dim UniqueValues As Collection
    Dim Value As Variant
    Dim ColumnIndex As Integer
    Dim InputBoxValue As String
    
    ' Ask the user to input the column to split
    InputBoxValue = InputBox("Enter the column letter you want to split by:")
    If InputBoxValue = "" Then Exit Sub ' If user cancelled the input box
    
    ' Convert column letter to number
    On Error Resume Next
    ColumnIndex = Range(InputBoxValue & "1").Column
    On Error GoTo 0
    Debug.Print "ColumnIndex: " & ColumnIndex
    
    If ColumnIndex = 0 Then
        MsgBox "Invalid column letter", vbExclamation
        Exit Sub
    End If
    
    Set OriginalWs = ThisWorkbook.ActiveSheet
    Set ColumnToSplit = OriginalWs.Columns(ColumnIndex).EntireColumn
    Debug.Print "ColumnToSplit.Address: " & ColumnToSplit.Address
    
    ' Use a collection to store unique values in the column
    Set UniqueValues = New Collection
    
    ' Try to add all cells' value to the collection. If it's already there, it will skip to the next one.
    On Error Resume Next
    For Each CurrentCell In ColumnToSplit.Cells
        If CurrentCell.Value <> "" Then
            Debug.Print "CurrentCell.Address: " & CurrentCell.Address & ", Type: " & TypeName(CurrentCell.Value)
        
            If IsArray(CurrentCell.Value) Then
                Debug.Print "CurrentCell.Value is an array"
            ElseIf IsError(CurrentCell.Value) Then
                Debug.Print "CurrentCell.Value is an error value"
            Else
                Debug.Print "CurrentCell.Value: " & CStr(CurrentCell.Value)
                
                ' Add the cell's value to the collection
                UniqueValues.Add CurrentCell.Value, CStr(CurrentCell.Value)
            End If
        End If
    Next CurrentCell

        On Error GoTo 0
        Debug.Print "UniqueValues.Count: " & UniqueValues.Count
    
    ' Loop through each unique value in the collection
    For Each Value In UniqueValues
        ' Open a new workbook
        Set NewWorkbook = Application.Workbooks.Add
        Set NewWs = NewWorkbook.ActiveSheet
        
        ' Copy the original sheet's headers to the new workbook
        OriginalWs.Rows(1).Copy Destination:=NewWs.Rows(1)
        
        ' Filter the original sheet based on the current unique value, and copy the visible (filtered) cells to the new workbook
        OriginalWs.Rows(1).AutoFilter Field:=ColumnIndex, Criteria1:=Value
        OriginalWs.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=NewWs.Cells(2, 1)
        
        ' Turn off filtering in the original sheet
        OriginalWs.AutoFilterMode = False
        
        ' Save and close the new workbook, using the unique value in the filename
        With NewWorkbook
            .SaveAs ThisWorkbook.Path & "\" & Value & ".xlsx"
            .Close
        End With
    Next Value
End Sub

