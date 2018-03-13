Option Explicit

Function get_folder_path() As String
'Date:          2018-03-09
'Summary:       opens filedialog and allows
'Credit:        derek from https://stackoverflow.com/a/23354219 (accessed 2018-03-05)
'               chris neilsen from https://stackoverflow.com/a/5975453 (accessed 2018-03-05)

'Declare variables
Dim diaFolder As FileDialog
Dim Msg As String, Title As String
Dim resStr As String
Dim response As Variant

'Open the file dialog
On Error GoTo ErrorHandler
Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)

'File dialog parameters
With diaFolder
    .AllowMultiSelect = False
    .Title = "Select a folder"
    .Show
    'Range("IC_Files_Path").Value2 = .SelectedItems(1)
End With

'Full path of folder as string
resStr = diaFolder.SelectedItems(1)

'Guarantee backslash at the end of path
If Right(resStr, 1) <> "\" Then resStr = resStr & "\"

'Return folderpath
get_folder_path = resStr
Set diaFolder = Nothing

Exit Function

ErrorHandler:
Msg = "No folder was selected. Process aborted."
Title = "No folder selected"
response = MsgBox(Msg, vbError, Title)

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Example how to use the function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub write_folder_path()

'Declare variables
Dim wb As Workbook
Dim ws As Worksheet
Dim path As String

Set wb = ActiveWorkbook
Set ws = wb.Sheets("Table1")

'Call the function
path = get_folder_path()

'Update path if selection was made
If path <> "" Then Range("B2").Value2 = get_folder_path()

End Sub
