Option Explicit

'###############################################################################

Sub import_multiple_csv_files()
'Date:          2018-03-08
'Summary:       imports data from multiple csv files into (a) specified table(s)
'               and saves the original files in another folder (archive)
'
'Functions:     write_string_to_cell: to create a list of already imported files
'               check_col_for_string: to check said list if the file has been imported before
'               
'Credit:        brettdf from https://stackoverflow.com/a/10382861 (accessed 2018-03-05)
'               Ron de Bruin from https://www.rondebruin.nl/win/s3/win026.htm (accessed 2018-03-05)

'Declare variables
Dim fPath   As String
Dim hPath   As String
Dim fName   As String
Dim fCSV    As String
Dim wbMST   As Workbook
Dim wbCSV   As Workbook
Dim wsCSV   As Worksheet
Dim lrMST   As Long
Dim lrCSV   As Long
Dim i       As Long
Dim t       As Single

'Makrotimer: starttime
t = Timer

'Change settings to speedup the code
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'No error messages are displayed, the default answers are taken
Application.DisplayAlerts = False

'Declare Sheets
Set wbMST = ThisWorkbook
Set wsMST = Sheets("Table2")

'Full path to source csv-files as string
fPath = "C:\Users\John\Source\"
If Right(fPath, 1) <> "\" Then fPath = fPath & "\"

'Full path where the csv-files should be archived as string
hPath = "C:\Users\John\Archive\"
If Right(hPath, 1) <> "\" Then hPath = hPath & "\"

'Naming convention of the csv-files, 
'if all csv-files in the folder should be used set to "*.csv"
fName = "filename??????.csv"

'Combine full path and naming convention
fCSV = Dir(fPath & filename)

'Counter
i = 0

'Exit macro, if no file found.
If fCSV = "" Then
    MsgBox "No csv-file with the given naming convention found." & vbNewLine & vbNewLine & _
           "Process aborted.", vbExclamation
    Exit Sub
End If

'Skip files with problems
On Error Resume Next

'Start Loop
Do While Len(fCSV) > 0
    
    'Check wheter file has been imported before, skip if True
    If check_col_for_string(fCSV, "Table1", "A") = False Then
        
        'Open csv-file with Local Excel settings (standard seperators for csv may vary)
        Set wbCSV = Workbooks.Open(fPath & fCSV, Local:=True) 
        
        'Declare sheet
        Set wsCSV = wbCSV.Sheets(1)
        
        'Get first empty row from column A in destination sheet (MAY NEED ADJUSTING)
        lrMST = wsMST.Cells(wsMST.Rows.Count, "A").End(xlUp).Row + 1
        
        'Get last populated row from column A in csv-source sheet (MAY NEED ADJUSTING)
        lrCSV = wsCSV.Cells(wsCSV.Rows.Count, "A").End(xlUp).Row
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Code to copy relevant data to destination sheet goes below here.
        'In the example, only specific colums are copied.
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~        
        
        ' copy columns B, and F to wsMST
        With wsMST
            .Range("A" & lrMST, "A" & (lrMST + lrCSV - 1)).Value2 = wsCSV.Range("B2:B" & lrCSV).Value2
            .Range("B" & lrMST, "B" & (lrMST + lrCSV - 1)).Value2 = wsCSV.Range("F2:F" & lrCSV).Value2
        End With

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~     
        
        'Close csv-file without saving changes
        wbCSV.Close False
        
        'Write csv-filename to the last row of "Table1", column 1
        Call write_string_to_cell(fCSV, "Table1", 1)
        
        'Increase counter
        i = i + 1
        
        'Copy original csv-file to archive
        FileCopy fPath & fCSV, hPath & fCSV
        
    End If
    
    'Prepare next csv-file 
    fCSV = Dir
    
Loop

'Clear up
Set wbCSV = Nothing

'Revert settings to speed up the code
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

'Status report
MsgBox "Count of files imported: " & i & vbNewLine & _
       "Time for code to run: " & Timer - t

End Sub

'###############################################################################

Function write_string_to_cell(str As String, ws As String, column As Long)
'Date:          2018-03-08
'Summary:       write string to the last row of a column and the time/date it was
'               done next to it
'Credit:        -
'Variables:
'@str           string to be written, e.g. "Hi"
'@worksheet     worksheet to be written into as string, e.g.  "Table1"
'@column        column of @worksheet as a number, e.g. for column A = 1, C = 3, etc

'Declare variables
Dim wb As Workbook
Dim ws As Worksheet
Dim lr As Long

'Declare sheets
Set wb = ThisWorkbook
Set ws = wb.Sheets(ws)

'Find first empty row in column
lr = ws.Cells(ws.Rows.Count, column).End(xlUp).Row + 1

'Write string to first empty row and import date/time next to it
ws.Cells(lr, column).Value2 = str
ws.Cells(lr, column).Offset(0, 1).Value2 = Format(Now(), "dd/mm/yyyy, h:m:s")

End Function

'###############################################################################

Function check_col_for_string(findstring As String, ws As String, col As String) As Boolean
'Date:          2018-03-08
'Summary:       check if a string is in a column, returns True / False
'Credit:        scott from https://stackoverflow.com/a/12643082 (accessed 2018-03-05)
'               Rene de la garza from https://stackoverflow.com/a/6265063 (accessed 2018-03-05)
'Variables:
'@findstring    string to to look for, e.g. "Hi"
'@worksheet     worksheet to look in, e.g.  "Table1"
'@column        column of @worksheet as a string, e.g. "A" or "B"
'Special cases:
'1) leading/trailling spaces are removed, so " " cannot be looked for (always returns False)

'Declare variables
Dim rng As Range
Dim ws as Worksheet

Set ws = Sheets(ws)

'Check if a string remains after trimming leading/trailling spaces
If Len(Trim(findstring)) > 0 Then
    
    'Search whole column
    With ws.Range(col & ":" & col)
        
        Set rng = .Find(What:=findstring, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
                        
        If Not rng Is Nothing Then
            check_col_for_string = True
        Else
            check_col_for_string = False
        End If
        
    End With
    
Else
    check_col_for_string = False
End If

End Function
