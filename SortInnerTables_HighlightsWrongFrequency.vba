Sub SortInnerTables_HighlightsWrongFrequency()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim tbl As Table
    Dim i As Integer
    Dim innerTbl As Table
    Dim cell As cell

    ' Loop through all tables in the document
    For Each tbl In doc.Tables
    tbl.Sort SortOrder:=wdSortOrderAscending, FieldNumber:=1
        For i = tbl.Rows.Count To 1 Step -1
            Set cell = tbl.cell(i, 8)
            ' Check if the 7th column contains an inner table
            If cell.Tables.Count > 0 Then
                Set innerTbl = cell.Tables(1)
                ' Sort the inner table by the 3rd column
                innerTbl.Sort SortOrder:=wdSortOrderAscending, FieldNumber:=1
                ' Delete half of the rows in the inner table
                
                For x = 1 To innerTbl.Rows.Count
                  With innerTbl.cell(x, 3)
                    y = Split(.Range.text, vbCr)(0)
                    If IsNumeric(y) Then
                      isValueValid = False
                      Select Case y
                      Case 3500, 750, 2350, 900, 2600, 1800, 2100, 26000, 0
                          isValueValid = True
                      End Select
                      If Not isValueValid Then
                          .Shading.BackgroundPatternColor = RGB(254, 140, 140)
                      End If
                    End If
                  End With
                Next
                
            End If
            
        Next i
    Next tbl
End Sub
