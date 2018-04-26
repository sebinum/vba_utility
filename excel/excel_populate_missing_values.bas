Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Step 0: create some dummy data (just for demonstration purposes)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub demo_dummy_values()

Dim wks As Worksheet
Dim rng As Range
Dim cell As Variant
Dim counter As Long

'set worksheet and range
Set wks = ThisWorkbook.Sheets(1)
Set rng = wks.Range("A1:C20")

'create some data
counter = 1
With wks
    For Each cell In rng
        cell.Value2 = counter
        counter = counter + 1
    Next cell
    'delete some, so there are missing values
    .Range("A1:C1").Value2 = ""
    .Range("A5:C8").Value2 = ""
    .Range("A19:C20").Value2 = ""
End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Step 1a: Call populate_missing_values with default "linear" argument
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub demo_populate_a()

Dim wks As Worksheet
Dim rng As Range

'set worksheet and range
Set wks = ThisWorkbook.Sheets(1)
Set rng = wks.Range("A1:C20")

Call populate_missing_values(rng)

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Step 1b: Call populate_missing_values with default "linear" argument
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub demo_populate_b()

Dim wks As Worksheet
Dim rng As Range

'set worksheet and range
Set wks = ThisWorkbook.Sheets(1)
Set rng = wks.Range("A1:C20")

Call demo_dummy_values 'reset data
Call populate_missing_values(rng, "static")

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'===============================================================================
'Summary:   populates missing values in the specified range by either
'           interpolating the missing values or statically write last seen value
'           to the empty cell. no extrapolation is used at the start/end of the
'           series if it starts / ends with missing values.
'           IMPORTANT: this procesdure will overwrite all of the specified range
'               with values, i.e. formulas will be converted to values
'Credit:    -
'Arguments:
'@my_range
'   the range of series in which to populate the missing values as range, e.g.
'   ThisWorkbook.Range("A2:B20")
'@my_method OPTIONAL
'   pick whether missing values should be linearly interpolated between known
'   values or statically updated with the last observed value as string, e.g.
'   "linear" for linear interpolation (default)
'   literally anything else but the default for static, such as "static" or "n"
'@my_by_row OPTIONAL
'   if the series are vertical (by rows) - True (default) or
'   if the series are horizontal (by colums) - False
'Changes------------------------------------------------------------------------
'Date       Change
'2018-04-26 written
'===============================================================================
Sub populate_missing_values( _
    my_range As Range, _
    Optional my_method As String = "linear", _
    Optional my_by_row As Boolean = True _
)

Dim arr_data() As Variant
Dim k_series As Long, k As Long
Dim i_observ As Long, i As Long, i_empty As Long
Dim arr_i_start() As Long
Dim lng_iplen As Long, j As Long
Dim dbl_upper As Double, dbl_lower As Double, dbl_incre As Double
Dim bln_method As Boolean

'determine if method is "linear"
bln_method = (my_method = "linear")

'load range into array and transpose if not by row
If my_by_row Then
    arr_data = my_range.Value2
Else
    arr_data = Application.WorksheetFunction.Transpose(my_range.Value2)
End If

'dimension of the array
k_series = UBound(arr_data, 2)
i_observ = UBound(arr_data, 1)

'array in which starting position for i is stored, since no extrapolation is
'done when a series begins with missing values / NULL
ReDim arr_i_start(1 To k_series)

'determine the starting position for each series k
For k = 1 To k_series
    'initilize / make sure is 0 before next element of k starts processing
    i_empty = 0
    For i = 1 To i_observ
        If Len(arr_data(i, k)) = 0 Then
            i_empty = i_empty + 1
        Else
            i_empty = i_empty + 1
            arr_i_start(k) = i_empty
            Exit For
        End If
    Next i
Next k

'populate values for k series but start @arr_i_start
For k = 1 To k_series
    'initilize / make sure is 0 before next element of k starts processing
    i_empty = 0
    
    For i = arr_i_start(k) To i_observ
        If Len(arr_data(i, k)) = 0 Then
            i_empty = i_empty + 1
        'if length is > 0 and empty counter is > the previous i_empty values
        'need to be populated
        ElseIf Len(arr_data(i, k)) > 0 And i_empty > 0 Then
            lng_iplen = i_empty + 1
            dbl_upper = arr_data(i, k)
            dbl_lower = arr_data(i - lng_iplen, k)
            
            'if default (linear) hasn't been changed do linear interpolation
            If bln_method Then
                'value by which each empty cell is incremented
                dbl_incre = (dbl_upper - dbl_lower) / lng_iplen
                
                'populate empty cells / missing values
                For j = 1 To i_empty
                    arr_data(i - j, k) = dbl_upper - (dbl_incre * j)
                Next j
            'else pick last known value and fill empty cells with it
            Else
                For j = 1 To i_empty
                    arr_data(i - j, k) = dbl_lower
                Next j
            End If
            'reset empty cell counter
            i_empty = 0
        End If
    Next i
Next k

'update range with filled values
If my_by_row Then
    my_range.Value2 = arr_data
Else
    my_range.Value2 = Application.WorksheetFunction.Transpose(arr_data)
End If

End Sub
