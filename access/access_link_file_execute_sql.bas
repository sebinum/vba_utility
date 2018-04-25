Option Compare Database
Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Step 1: Create and save a link/import specification for the file(s) used
'        In the demo it is called acLink_dailysales.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Step 2: Add sql_selects that should be run to link_file_execute_sql
'        search for /start/ to find where to place them and an example
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Step 3: Call link_file_execute_sql with parameters
Sub demo_file_link()

'Demo A: without checking for already imported files
Call link_file_execute_sql( _
                           my_path:="C:\Users\John\Archive\", _
                           my_file:="dailysales_??????.csv", _
                           my_specification_name:="acLink_dailysales" _
                           )

'Demo B: with checking for already imported files in a table and updating
'        that table
Call link_file_execute_sql( _
                           my_path:="C:\Users\John\Archive\", _
                           my_file:="dailysales_??????.csv", _
                           my_specification_name:="acLink_dailysales", _
                           my_imported_files_table:="tblProcessedFiles", _
                           my_imported_files_field:="filename" _
                           )
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'===============================================================================
'Summary:   counts the files fullfilling the name criteria in a folder and
'           returns the result as long
'Credit:    Ryan Wells from (accessed 2018-04-19)
'           https://wellsr.com/vba/2016/excel/vba-count-files-in-folder/
'           Bathsheba from https://stackoverflow.com/a/16612000
'               (accessed 2018-04-19)
'Arguments:
'@my_path
'   the path to the folder the files lie in, e.g. "C:\John\Reports\"
'@my_filename
'   filename or type, e.g. "report??????.csv", "*.csv", "*.txt", "*.xlsx"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-19 written
'===============================================================================
Private Function count_files_in_folder( _
    ByVal my_path As String, _
    ByVal my_filename As String _
    ) As Long
'declare variables
Dim str_path As String
Dim var_file As Variant
Dim i As Integer

'check for backslash and add if necessary
If Right(my_path, 1) <> "\" Then str_path = my_path & "\" Else str_path = my_path

'initialize full path
var_file = Dir(str_path & my_filename)

'count files
While (var_file <> "")
    i = i + 1
    var_file = Dir
Wend

count_files_in_folder = i

End Function

'===============================================================================
'Summary:   check whether a table with a given name exists in the database,
'           returns true/false
'Credit:    KevenDenen from https://stackoverflow.com/a/2985861
'               (accessed 2017-08-02)
'Arguments:
'@table_name
'   table name to check for, e.g. "tblCurrentReportxx"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-19 written
'===============================================================================
Private Function check_db_for_table(my_table_name As String) As Boolean

Dim dao_db As DAO.Database
Dim dao_td As DAO.TableDef

Set dao_db = CurrentDb
On Error Resume Next
Set dao_td = dao_db.TableDefs(my_table_name)

'return True if found, False for an error
check_db_for_table = (Err.Number = 0)

Err.Clear
    
End Function

'===============================================================================
'Summary:   update the status bar with a given string, e.g. useful in looping
'           through multiple huge files and performing a task
'           should be called with 'Call Status("")' after loop is finished to
'           set back to default
'Credit:    -
'Arguments:
'@status_text
'   the text to be display, e.g. "Processed file " & lng_file & " from " & '_'
'                                lng_filecount & "."
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-18 written
'===============================================================================
Private Sub update_system_status(my_status_text As String)
    
Dim var_status As Variant
    
If my_status_text = "" Then
    'if = "" clear status (default)
    var_status = SysCmd(acSysCmdClearStatus)
Else
    'else display input string
    var_status = SysCmd(acSysCmdSetStatus, my_status_text)
End If

End Sub

'===============================================================================
'Summary:   executes SQL-code from a string, will fail on error and tell you why
'Credit:    Danny Lesandrini from (accessed 2018-04-19)
'           https://www.databasejournal.com/features/msaccess/article.php/
'           10895_3505836_2/Executing-SQL-Statements-in-VBA-Code.htm
'Arguments:
'@my_sql_statement
'   the sql statement which should be executed, e.g. "SELECT * FROM tblA"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-19 written
'===============================================================================
Private Sub execute_sql(my_sql_statement As String)

Dim dao_db As DAO.Database

Set dao_db = CurrentDb

dao_db.Execute my_sql_statement, dbFailOnError

End Sub

'===============================================================================
'Summary:   returns a list of files in a folder in form of an array
'Credit:    Dave Rado from (accessed 2018-04-20)
'           https://wordmvp.com/FAQs/MacrosVBA/ReadFilesIntoArray.htm
'Arguments:
'@my_path
'   the path to the folder the files lie in, e.g. "C:\John\Reports\"
'@my_file
'   filename or type, e.g. "report??????.csv", "*.csv", "*.txt", "*.xlsx"
'@my_imported_files_table OPTIONAL
'   if files that have been processed already should be skipped, check a table
'   for a record that matches the name of the file
'@my_imported_files_field OPTIONAL
'   the field which is checked for a matching record in @my_imported_files_table
'@my_max_arr_size OPTIONAL
'   initial size of the array, only needs to be changed if over 1000 files are
'   in the folder
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-20 written
'===============================================================================
Private Function arr_list_of_files( _
    ByVal my_path As String, _
    ByVal my_file As String, _
    Optional ByVal my_imported_files_table As String, _
    Optional ByVal my_imported_files_field As String, _
    Optional ByVal my_max_arr_size As Long = 1000 _
    ) As Variant

Dim i As Long
Dim str_file As String
Dim lgn_match As Long
Dim bln_opt As Boolean

'create a dynamic array variable, and then declare its initial size
Dim arr_file_list() As String
ReDim arr_file_list(my_max_arr_size)

'check whether optional parameters where supplied
bln_opt = Len(my_imported_files_table) > 0 And Len(my_imported_files_field) > 0

'Loop through all the files in the directory
str_file = Dir$(my_path & my_file)
Do While str_file <> ""

    'if optional parameter has been set and file matches a record in specified
    'table, skip that file and continue with the next file
    If bln_opt Then
        'count the records in @my_imported_files_table that match the filename
        lgn_match = DCount(my_imported_files_field, my_imported_files_table, _
                           my_imported_files_field & "= '" & str_file & "'")
        
        'if one or more records match, initiliaze next file
        If lgn_match > 0 Then
            'initialize next file
            str_file = Dir$
        'else write the filename to the array before initiliazing the next file
        Else
            'write filename to array
            arr_file_list(i) = str_file

            'initialize next file
            str_file = Dir$
    
            'counter
            i = i + 1
        End If
    'kicks in if no optional parameters have been passed to the function
    Else
        'write filename to array
        arr_file_list(i) = str_file
              
        'initialize next file
        str_file = Dir$
    
        'counter
        i = i + 1
    End If
Loop

If i > 0 Then
    'reset the size of the array without losing its values
    ReDim Preserve arr_file_list(i - 1)

    'return array
    arr_list_of_files = arr_file_list
Else
    arr_list_of_files = Null
End If

End Function

'===============================================================================
'Summary:   links each file matching a given name to the database based on a
'           import specification, execute sql(s) and remove the linked tables.
'           this is especially useful for processing multiple large files that
'           are to big to import (tested with text files ~250MB in size)
'Credit     -
'Arguments:
'@my_path
'   the path to the folder the files lie in, e.g. "C:\John\Reports\"
'@my_file
'   filename or type, e.g. "report??????.csv", "*.csv", "*.txt", "*.xlsx"
'@my_spec_name
'   name of the import specification as string, e.g. "weeklyrep_is"
'   IMPORTANT: it is up to the user to determine an adequate specification and
'              sql selects that work with that specification
'@my_proc_files_table OPTIONAL
'   if files that have been processed already should be skipped, check a table
'   for a record that matches the name of the file
'@my_proc_files_field OPTIONAL
'   the field which is checked for a matching record in @my_proc_files_table
'
'ToDo:
'      insert the sql which should be executed in the indented block in this code
'   the variable str_table_name contains the name of the linked file and can be
'   used in the sql select. e.g. "FROM " & str_table_name & " AS a "
'       if the optional parameters are supplied, each processed filename will be
'   inserted in @my_proc_files_field. if further fields should be populated e.g.
'   time of processing, who processed it etc. the sql needs to be adjusted
'       if more than 1000 files are in the folder @my_path then the call to
'    arr_list_files needs adjusting (supply the optional argument with a higher
'    value
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-25 written
'===============================================================================
Sub link_file_execute_sql(ByVal my_path As String, _
                          ByVal my_file As String, _
                          ByVal my_spec_name As String, _
                          Optional ByVal my_proc_files_table As String, _
                          Optional ByVal my_proc_files_field As String)

Dim str_file As String
Dim str_table_name As String
Dim str_initial_status As String
Dim str_remaining_time As String
Dim var_file_arr As Variant
Dim var_arr_ele As Variant
Dim lng_filecount_all As Long
Dim lng_files_proc As Long
Dim sng_timer_beg As Single
Dim sng_timer_end As Single
Dim sng_timer_avg As Single
Dim i As Long
Dim bln_opt As Boolean

'Check whether slash at the end was input correctly
If Right(my_path, 1) <> "\" Then my_path = my_path & "\" Else my_path = my_path

'Only implemented for text files (txt and csv), will Exit procedure otherwise
If Right(my_file, 4) <> ".txt" And Right(my_file, 4) <> ".csv" Then
    MsgBox "The ending of a file name must be '.txt' or '.csv'." & vbNewLine & _
           vbNewLine & "Process aborted.", vbExclamation
    Exit Sub
End If

str_initial_status = "Processing first file and calculating " & _
                     "average procedure runtime..."

'determines whether strings have been passed to both optional parameters
bln_opt = Len(my_proc_files_table) > 0 And Len(my_proc_files_field) > 0

'error handler for empty array and its consequences
On Error GoTo exit_with_error

'count the number of files that match @my_file to calculate approximate procedure
'runtime and create a nice status update
lng_filecount_all = count_files_in_folder(my_path, my_file)

'list of the relevant files as an array
var_file_arr = arr_list_of_files(my_path, my_file, my_proc_files_table, _
                                 my_proc_files_field)

'count the number of relevant files to be processed
lng_files_proc = UBound(var_file_arr) + 1


For Each var_arr_ele In var_file_arr
    'start time processing a file
    sng_timer_beg = Timer
    
    'filename without extension
    str_table_name = Left(var_arr_ele, Len(var_arr_ele) - 4)
    
    'iterator
    i = i + 1
    
    'update the status
    If i = 1 Then
        Call update_system_status(str_initial_status)
    Else
       'average time it took to process files
        sng_timer_avg = _
            ((sng_timer_avg * (i - 2)) + sng_timer_end) / (i - 1)
        
        'estimate of remaining time
        str_remaining_time = _
            Format((sng_timer_avg * (lng_files_proc - i + 1)) / 86400, _
            "hh:mm:ss")
            
        Call update_system_status( _
            "Processing file " & i & "/" & lng_files_proc & " (" & _
            var_arr_ele & "). Estimated time to process the remaining " & _
            "file(s): " & str_remaining_time)
    End If
    
    'link table to database
    DoCmd.TransferText TransferType:=acLinkDelim, _
                       SpecificationName:=my_spec_name, _
                       TableName:=str_table_name, _
                       FileName:=my_path & var_arr_ele, _
                       HasFieldNames:=True, _
                       HTMLTableName:=""
        
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '/start/ code, e.g. run some SQL-Selects, goes here.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    'made up demo sql
    Call execute_sql( _
        "INSERT INTO [tblSales] ( Customer, ID, Product, Revenue )" & _
        "SELECT a.Customer, a.ID, a.Product, a.Revenue" & _
        "FROM " & str_table_name & " AS a" & _
        "WHERE ((a.Customer) In ('Smith', 'Rogers'))" & _
        "GROUP BY a.Customer;" _
        )
    
    'will only run if strings have been passed to both optional parameters
    'this block can be further extended to include e.g. time of calling this sub
    If bln_opt Then
        Call execute_sql("INSERT INTO [" & my_proc_files_table & "]" & _
                         "( " & my_proc_files_field & " ) VALUES " & _
                         "('" & var_arr_ele & "');")      
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '/end/ code, e.g. run some SQL-Selects, goes here.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    'remove link to file
    If check_db_for_table(str_table_name) Then
        DoCmd.DeleteObject acTable, str_table_name
    End If
    
    'end time processing a file
    sng_timer_end = Timer - sng_timer_beg
Next var_arr_ele

'reset system status to default
Call update_system_status("")

'tell me what happened
MsgBox i & " files matching the namingconvention (" & my_file & ") haven been " & _
       "processed successfully." & vbNewLine & vbNewLine & _
       (lng_filecount_all - i) & " files were skipped due to the optional " & _
       "parameterisation." & vbNewLine & vbNewLine & "The average " & _
       "processing-time per file was: " & _
       Format(sng_timer_avg / 86400, "hh:mm:ss"), _
       vbInformation, _
       "Files successfully processed"

Exit Sub

'error handler, kicks in if no files to process are found or if the optional
'parameters where populated and all valid files have been processed previously
exit_with_error:

If bln_opt And lng_filecount_all > 0 Then
    MsgBox lng_filecount_all & " file(s) match(es) the namingconvention (" & _
           my_file & ") and has/have been processed previously." & vbNewLine & _
           vbNewLine & "No further action was taken.", vbInformation
Else
    MsgBox "No file(s) matching the namingconvetion found." & vbNewLine & _
           vbNewLine & "Convention: '" & my_file & "'" & vbNewLine & _
           vbNewLine & "Process aborted.", vbExclamation
End If

End Sub
