Option Explicit

Function get_path(file_or_folder As String) As String
'Date:          2018-03-14
'Summary:       opens filedialog and allows selection of a file/folder,
'               the files'/folders' full path is returned as a string
'Credit:        derek from https://stackoverflow.com/a/23354219 (accessed 2018-03-05)
'               chris neilsen from https://stackoverflow.com/a/5975453 (accessed 2018-03-05)
' Update:       2018-03-21: added functionality to choose between file / folder picker

'Declare variables
Dim afd             As FileDialog
Dim str_msg         As String
Dim str_title       As String
Dim str_selection   As String
Dim var_response    As Variant

'Error Handler
On Error GoTo ErrorHandler

'If input is file, open FilePicker
If file_or_folder = "file" Then
    Set afd = Application.FileDialog(msoFileDialogFilePicker)

'If input is file, open FolderPicker
ElseIf file_or_folder = "folder" Then
    Set afd = Application.FileDialog(msoFileDialogFolderPicker)
'Else abort process and say why
Else
    MsgBox _
        "Only 'file' or 'folder' are valid entries for the function get_path()." & _
        vbNewLine & vbNewLine & _
        "Please check your input.", _
        vbExclamation + vbOKOnly, _
        "Function aborted"
    Exit Function
End If

'File dialog parameters (currently no multiselect for files / folders)
With afd
    .AllowMultiSelect = False
    .Title = "Select a " & file_or_folder
    .Show
End With

'Full path of file / folder as string
str_selection = afd.SelectedItems(1)

'Guarantee backslash at the end of folder-path
If file_or_folder = "folder" And Right(str_selection, 1) <> "\" Then
    str_selection = str_selection & "\"
End If

'Return folderpath as String
get_path = str_selection

'Clean up
Set afd = Nothing
On Error GoTo 0

Exit Function

'Error Handling
ErrorHandler:
str_msg = "No folder was selected. Process aborted."
str_title = "No folder selected"
var_response = MsgBox(str_msg, vbError, str_title)
On Error GoTo 0

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Examples how to use the function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub write_folder_path()

'Declare variables
Dim wb As Workbook
Dim ws As Worksheet
Dim path As String

Set wb = ActiveWorkbook
Set ws = wb.Sheets("Tabelle1")

'Call the function
path = get_path("folder")

'Update path if selection was made
If path <> "" Then Range("B2").Value2 = path

End Sub

Sub write_file_path()

'Declare variables
Dim wb As Workbook
Dim ws As Worksheet
Dim path As String

Set wb = ActiveWorkbook
Set ws = wb.Sheets("Tabelle1")

'Call the function
path = get_path("file")

'Update path if selection was made
If path <> "" Then Range("B2").Value2 = path

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
