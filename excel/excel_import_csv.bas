Option Explicit

'===============================================================================
'Summary:   write a files name, its' modified date and the time of using this
'           sub in the first empty row of a specified column
'Credit:    Siddharth Rout from https://stackoverflow.com/a/10823572
'               (accessed 2018-03-23)
'Arguments:
'@my_filepath
'   String to be written, e.g. "C:\Users\John\"
'@my_filename
'   full name of file to be written, e.g. "myfile1.csv"
'@my_worksheet
'   Worksheet to be written into as string, e.g.  "Table1"
'@my_column
'   Column of @my_worksheet as a String, e.g. "A", "B" or "Y"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-03-08 written
'2018-03-22 style changes implemented
'2018-03-23 added filemodified date/time and renamed sub
'===============================================================================
Sub write_fileinformation(my_filepath As String, _
                          my_filename As String, _
                          my_worksheet As String, _
                          my_column As String)

'Declare variables
Dim wks As Worksheet
Dim lng_lr As Long
Dim str_moddate As String
Dim str_usetime As String

'Set sheet
Set wks = ThisWorkbook.Sheets(my_worksheet)

'Full path to file
str_moddate = FileDateTime(my_filepath & my_filename)

'Format for the date/time displayed when using this sub
str_usetime = Format(Now(), "dd/mm/yyyy, hh:mm:ss")

With wks
    'Find first empty row in column
    lng_lr = .Cells(.Rows.Count, my_column).End(xlUp).Row + 1
    'Write string to first empty row and import date/time next to it
    .Cells(lng_lr, my_column).Value2 = my_filename
    .Cells(lng_lr, my_column).Offset(0, 1).Value2 = str_moddate
    .Cells(lng_lr, my_column).Offset(0, 2).Value2 = str_usetime
End With

End Sub

'===============================================================================
'Summary:   check if a string is in a my_columnumn, returns True / False
'           done next to it
'Credit:    scott from https://stackoverflow.com/a/12643082
'               (accessed 2018-03-05)
'           Rene de la garza from https://stackoverflow.com/a/6265063
'               (accessed 2018-03-05)
'Arguments:
'@my_string
'   String to to look for , e.g. "Hi"
'@my_worksheet
'   Worksheet to look in, e.g.  "Table1"
'@my_column
'   Column of @my_worksheet to look in, e.g. "A" or "B"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-03-08 written
'2018-03-22 style changes implemented
'===============================================================================
Function check_col_for_string(my_string As String, _
                              my_worksheet As String, _
                              my_column As String _
                              ) As Boolean
'Declare variables
Dim rng As Range
Dim wks As Worksheet

'Set sheet
Set wks = ThisWorkbook.Sheets(my_worksheet)

'Check if a string remains after trimming leading/trailling spaces
If Len(Trim(my_string)) > 0 Then
    
    'Search whole column
    With wks.Range(my_column & ":" & my_column)
        
        Set rng = .Find(What:=my_string, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        
        'True, if a match is found
        If Not rng Is Nothing Then
            check_col_for_string = True
        'False, if no match is found
        Else
            check_col_for_string = False
        End If
        
    End With
    
Else
    'False, if no string ("") remains after trimming
    check_col_for_string = False
End If

End Function

'===============================================================================
'Summary:   imports data from multiple csv files into (a) specified table(s)
'           of the workbook the code runs from and saves the original files in
'           another folder (archive)
'Credit:    brettdf from https://stackoverflow.com/a/10382861
'               (accessed 2018-03-05)
'           Ron de Bruin from https://www.rondebruin.nl/win/s3/win026.htm
'               (accessed 2018-03-05)
'Arguments:
'@my_destin_sht
'   The worksheet in which to import the data from the csv files, e.g. "Table1"
'@my_source_path
'   The full path to where the csv files are, e.g. "C:\User\John\Source"
'@my_archive_path
'   The full path to where the csv files are to be archived, e.g. "C:\Archive"
'@my_source_fname
'   The naming convention of the csv file, e.g. "*.csv" or "data_??????.csv"
'@my_imp_list_sht
'   The worksheet in which to write the imported csv-filenames, e.g. "Table2"
'@my_imp_list_col
'   The column of @my_imp_list_sht in which to write, e.g. "A"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-03-08 written
'2018-03-22 style changes implemented
'Planned changes----------------------------------------------------------------
'1) outsource the block where the data wrangling is done to another sub,
'   that way recylcling the code will be limited to changing one isolated sub
'   and calling import_csv
'===============================================================================
Sub import_csv_files(my_destin_sht As String, _
                     my_source_path As String, _
                     my_source_fname As String, _
                     my_archive_path As String, _
                     my_imp_list_sht As String, _
                     my_imp_list_col As String)
'Declare variables
Dim wkb_mst             As Workbook
Dim wkb_csv             As Workbook
Dim wks_mst             As Worksheet
Dim wks_csv             As Worksheet
Dim str_path_source     As String
Dim str_path_archive    As String
Dim str_filename        As String
Dim str_csv             As String
Dim lng_lr_mst          As Long
Dim lng_lr_csv          As Long
Dim lng_lr_new          As Long
Dim lng_counter         As Long
Dim sng_timer           As Single
Dim bln_imp             As Boolean

'Makrotimer: starttime
sng_timer = Timer

'Change settings to speedup the code / suppress display alerts
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .DisplayAlerts = False
End With

'Declare Sheets
Set wkb_mst = ThisWorkbook
Set wks_mst = wkb_mst.Sheets(my_destin_sht)

'Full path to source csv-files as string
str_path_source = my_source_path
If Right(str_path_source, 1) <> "\" Then
    str_path_source = str_path_source & "\"
End If

'Full path where the csv-files should be archived as string
str_path_archive = my_archive_path
If Right(str_path_archive, 1) <> "\" Then
    str_path_archive = str_path_archive & "\"
End If

'Naming convention of the csv-files,
str_filename = my_source_fname

'Combine full path and naming convention
str_csv = Dir(str_path_source & my_source_fname)

'Counter
lng_counter = 0

'Exit macro, if no file found.
If str_csv = "" Then
    MsgBox "No csv-file with the given naming convention found." & _
            vbNewLine & vbNewLine & _
           "Process aborted.", vbExclamation
    Exit Sub
End If

'Skip files with problems (BETTER WAY TO HANDLE THIS NEEDED)
On Error Resume Next

'Start Loop
Do While Len(str_csv) > 0
    
    'Check wheter file has been imported before
    bln_imp = check_col_for_string(str_csv, my_imp_list_sht, my_imp_list_col)
    
    
    'If file hasn't been imported, do something
    If bln_imp = False Then
        
        'Open csv-file with Local Excel settings
        Set wkb_csv = Workbooks.Open(str_path_source & str_csv, Local:=True)
        
        'Declare sheet
        Set wks_csv = wkb_csv.Sheets(1)
        
        'Get first empty row from column A in destination sheet
        With wks_mst
            lng_lr_mst = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
        'Get last row from column A in csv-source sheet
        With wks_csv
            lng_lr_csv = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Code to copy relevant data to destination sheet goes below here.
        'In the example, only specific colums are copied.
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        'new end row, where to copy data
        lng_lr_new = lng_lr_mst + lng_lr_csv - 2
        
        ' copy columns B, and F to wks_mst
        With wks_mst
            .Range("A" & lng_lr_mst, "A" & lng_lr_new).Value2 = _
                wks_csv.Range("B2:B" & lng_lr_csv).Value2
            .Range("B" & lng_lr_mst, "B" & lng_lr_new).Value2 = _
                wks_csv.Range("F2:F" & lng_lr_csv).Value2
        End With

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        'Close csv-file without saving changes
        wkb_csv.Close False
        
        'Write fileinformation to @my_imp_list_sht
        Call write_fileinformation(str_path_source, str_csv, _
                                   my_imp_list_sht, my_imp_list_col)
        
        'Increase counter
        lng_counter = lng_counter + 1
        
        'Copy original csv-file to archive
        FileCopy str_path_source & str_csv, str_path_archive & str_csv
        
    End If
    
    'Prepare next csv-file
    str_csv = Dir
    
Loop

'Clear up
Set wkb_csv = Nothing

'Revert settings to speed up the code
With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .DisplayAlerts = True
End With

'Status report
MsgBox "Count of files imported: " & lng_counter & vbNewLine & _
       "Time for code to run: " & Timer - sng_timer

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Example how to use the import_csv_files /start
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub testimport()

Call import_csv_files("Table2", _
                      "C:\sourcefolder\", _
                      "U:\archivefolder\", _
                      "somedailyreport_??????.csv", _
                      "Table1", "A")
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Example how to use the import_csv_files /end
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
