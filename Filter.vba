Sub Filter()
    On Error Resume Next
    
    Dim wsRes As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim filterValues As Variant
    Dim wsSheet1 As Worksheet
    Dim selectedRange As Range
    Dim cell As Range
    Dim filterCount As Integer
    Dim unrelatedCount As Integer
    Dim found As Boolean
    Dim filteredSelected As String
    Dim unrelatedItems As String
    Dim selectedWithoutRelation As String
    
    Set wsRes = ThisWorkbook.Worksheets("TD Summary")
    Set wsSheet1 = ThisWorkbook.Worksheets("Warehouse Report")
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells in Sheet1 before running the filter."
        Exit Sub
    End If
    
    Set selectedRange = Selection
    ReDim filterValues(1 To selectedRange.Cells.Count)
    Dim i As Integer
    i = 1
    For Each cell In selectedRange
        filterValues(i) = cell.Value
        i = i + 1
    Next cell
    
    If UBound(filterValues) = 0 Then
        MsgBox "Please select at least one cell before running the filter."
        Exit Sub
    End If
    
    Set pt = wsRes.PivotTables("Table1")
    Set pf = pt.PivotFields("sap")

    filterCount = 0
    unrelatedCount = 0
    filteredSelected = ""
    unrelatedItems = ""
    selectedWithoutRelation = ""
    
    If Application.WorksheetFunction.CountA(selectedRange) = 0 Then
        MsgBox "Please select cells with data to filter."
        Exit Sub
    End If
    
    If pt.PivotFields("sap").Orientation = xlHidden Then
        MsgBox "The field 'sap' is not present in the pivot table 'Table1'. Verify that the field name is correct."
        Exit Sub
    End If
    
    pf.ClearAllFilters
    
    For Each Item In pf.PivotItems
        found = False
        For Each cell In selectedRange
            If cell.Value = Item.Value Then
                found = True
                Exit For
            End If
        Next cell
        If found Then
            Item.Visible = True
            filterCount = filterCount + 1
            If filterCount > 0 Then
                If filteredSelected = "" Then
                    filteredSelected = Item.Value
                Else
                    filteredSelected = filteredSelected & ", " & Item.Value
                End If
            End If
        Else
            Item.Visible = False
            If unrelatedItems = "" Then
                unrelatedItems = Item.Value
            Else
                unrelatedItems = unrelatedItems & ", " & Item.Value
            End If
            unrelatedCount = unrelatedCount + 1
        End If
    Next Item
    
    pt.RefreshTable
    
    wsRes.Activate
    
    For Each cell In selectedRange
        found = False
        For Each Item In pf.PivotItems
            If cell.Value = Item.Value Then
                found = True
                Exit For
            End If
        Next Item
        If Not found Then
            unrelatedCount = unrelatedCount + 1
            If selectedWithoutRelation = "" Then
                selectedWithoutRelation = cell.Value
            Else
                selectedWithoutRelation = selectedWithoutRelation & ", " & cell.Value
            End If
        End If
    Next cell
    
    If filterCount > 0 Then
        If unrelatedCount > 0 Then
            MsgBox "Filtered " & filterCount & " SAP items: " & filteredSelected & ". " & _
                   vbCrLf & vbCrLf & _
                   "The following selected items have no relation to SAP: " & selectedWithoutRelation & "."
        Else
            MsgBox "Filtered " & filterCount & " SAP items: " & filteredSelected & ". " & _
                   vbCrLf & vbCrLf & _
                   "No items found that have no relation to SAP." & _
                   vbCrLf & vbCrLf & _
                   "The following selected items have no relation to SAP: " & selectedWithoutRelation & "."
        End If
    Else
        If unrelatedCount > 0 Then
            MsgBox "No SAP items filtered. " & _
                   vbCrLf & vbCrLf & _
                   "The following selected items have no relation to SAP: " & selectedWithoutRelation  & "."
        Else
            MsgBox "No SAP items filtered. " & _
                   vbCrLf & vbCrLf & _
                   "No items found that have no relation to SAP." & _
                   vbCrLf & vbCrLf & _
                   "The following selected items have no relation to SAP: " & selectedWithoutRelation & "."
        End If
    End If
End Sub