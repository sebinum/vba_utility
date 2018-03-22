Option Explicit

'===============================================================================
'Summary:   opens filedialog and allows selection of a file/folder,
'           the files'/folders' full path is returned as a string
'Credit:    derek from https://stackoverflow.com/a/23354219
'               (accessed 2018-03-05)
'           chris neilsen from https://stackoverflow.com/a/5975453
'               (accessed 2018-03-05)
'Arguments:
'file_or_folder
'   decide whetere the a files' or a folders' full str_path should be picked and
'   returned. valid entries are "file" and "folder"
'Changes------------------------------------------------------------------------
'Date       Change
'2018-03-14 written
'2018-03-21 added functionality to choose between file / folder picker
'2018-03-22 style changes implemented
'===============================================================================
Function get_path(file_or_folder As String) As String

'Declare variables
Dim fdo             As FileDialog
Dim str_msg         As String
Dim str_title       As String
Dim str_selection   As String
Dim var_response    As Variant

'Error Handler
On Error GoTo ErrorHandler

'If input is file, open FilePicker
If file_or_folder = "file" Then
    Set fdo = Application.FileDialog(msoFileDialogFilePicker)

'If input is folder, open FolderPicker
ElseIf file_or_folder = "folder" Then
    Set fdo = Application.FileDialog(msoFileDialogFolderPicker)

'Else abort process and say why
Else
    MsgBox _
        "Only 'file' or 'folder' are valid entries for get_path()." & _
        vbNewLine & vbNewLine & _
        "Please check your input.", _
        vbExclamation + vbOKOnly, _
        "Function aborted"
    Exit Function
End If

'File dialog parameters (currently no multiselect for files / folders)
With fdo
    .AllowMultiSelect = False
    .Title = "Select a " & file_or_folder
    .Show
End With

'Full str_path of file / folder as string
str_selection = fdo.SelectedItems(1)

'Guarantee backslash at the end of folder-path
If file_or_folder = "folder" And Right(str_selection, 1) <> "\" Then
    str_selection = str_selection & "\"
End If

'Return folderpath as String
get_path = str_selection

'Clean up
Set fdo = Nothing
On Error GoTo 0

Exit Function

'Error Handling
ErrorHandler:
MsgBox "No folder was selected. Process aborted.", _
       vbError, _
       "No " & file_or_folder & " selected."
'Reset ErrorHandler
On Error GoTo 0

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Examples how to use the function /start
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub write_folder_path()

'Declare variables
Dim wks         As Worksheet
Dim str_path    As String

Set wks = ThisWorkbook.Sheets("Tabelle1")

'Call the function
str_path = get_path("folder")

'If selection was made, write path to cell A2
If str_path <> "" Then Range("A2").Value2 = str_path

End Sub

Sub write_file_path()

'Declare variables
Dim wks         As Worksheet
Dim str_path    As String

Set wks = ThisWorkbook.Sheets("Tabelle1")

'Call the function
str_path = get_path("file")

'If selection was made, write path to cell A2
If str_path <> "" Then Range("A2").Value2 = str_path

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Examples how to use the function /end
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
