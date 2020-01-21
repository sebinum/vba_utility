Option Explicit

'===============================================================================
'Description:   search for the last populated row in a specified column
'               if no column is specified, it defaults to column A / 1
'Credit:    
'   Chris Newman from https://www.thespreadsheetguru.com/ (accessed 2020-01-09)
'Arguments:
'@wks_worksheet_name
'   worksheet to look on as an worksheet object, e.g. Sheets("Table1")
'@var_column_index [optional] -> default: column A / 1
'   Column of @wks_worksheet_name as a Variant to allow for entry of column
'   in R1C1 - e.g. 1, 2, 3 (Long) - or A1 - e.g. "A", "B", "C" (String) - format
'Changes------------------------------------------------------------------------
'Date       Change
'2020-01-09 written
'===============================================================================
Function last_row(wks_worksheet_name As Worksheet, _
                  Optional var_column_index As Variant = 1) As Long
    
    With wks_worksheet_name
        last_row = .Cells(.Rows.Count, var_column_index).End(xlUp).Row
    End With

End Function

'===============================================================================
'Description:   search for the last populated row in a specified row
'           if no row is specified, it defaults to row 1
'Credit:    
'   Chris Newman from https://www.thespreadsheetguru.com/ (accessed 2020-01-09)
'Arguments:
'@wks_worksheet_name
'   worksheet to look on as an worksheet object, e.g. Sheets("Table1")
'@lng_row_index [optional] -> default: row 1
'   Column of @wks_worksheet_name as a String or Long, e.g. "A", "C" or 1, 3
'Changes------------------------------------------------------------------------
'Date       Change
'2020-01-09 written
'===============================================================================
Function last_column(wks_worksheet_name As Worksheet, _
                     Optional lng_row_index As Long = 1) As Variant
    
    With wks_worksheet_name
        last_column = .Cells(lng_row_index, .Columns.Count).End(xlToLeft).Column
    End With

End Function

'===============================================================================
'Description:
'   checks whether a filter is active on the specified worksheet and whether
'   the data is filtered, if both applies returns True, otherwise False
'Credit:
'   Batman from https://www.ozgrid.com/forum/index.php?thread/56458-reset-all-filters-to-all/&postID=713635#post713635
'   (accessed 2020-01-09)
'Arguments:
'@wks_worksheet_name
'   worksheet to reset filter on as an worksheet object, e.g. Sheets("Table1")
'Changes------------------------------------------------------------------------
'Date       Change
'2020-01-09 written
'===============================================================================
Function data_filtered(wks_worksheet_name As Worksheet) As Boolean

    With wks_worksheet_name
        If .AutoFilterMode Then
            If .FilterMode Then
                data_filtered = True
            Else
                data_filtered = False
            End If
        End If
    End With

End Function

'===============================================================================
'Description:
'   checks whether a worksheet with a specified name exists
'Credit:
'   Dante May Code  from stackoverflow and Tim Williams from Stackoverflow
'   (accessed 2020-01-17)
'Arguments:
'@str_worksheet_name
'   worksheet name to look for as a String, e.g. "Table1" or "income"
'@wkb_workbook [optional] -> default: ThisWorkbook
'   Excel-Workbook to search as Workbook-Object, e.g. ThisWorkbook
'Changes------------------------------------------------------------------------
'Date       Change
'2020-01-17 written
'===============================================================================
Function worksheet_exists(str_worksheet_name As String, _
                          Optional wkb_workbook As Workbook) As Boolean
    
    Dim wks_worksheet As Worksheet

    'defaults to ThisWorkbook (the workbook the procedure is called from)
    'if no parameter is passed
    If wkb_workbook Is Nothing Then Set wkb_workbook = ThisWorkbook
    
    For Each wks_worksheet In wkb_workbook.Worksheets
        If str_worksheet_name = wks_worksheet.Name Then
            worksheet_exists = True
            Exit Function
        End If
    Next wks_worksheet

End Function
