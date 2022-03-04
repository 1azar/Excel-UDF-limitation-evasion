Attribute VB_Name = "Test_fw"
Public Function some_func_returns_array()
    Dim my_arr() As Variant
    ReDim my_arr(0 To 2, 0 To 2)
    For i = 0 To UBound(my_arr, 1)
        For j = 0 To UBound(my_arr, 2)
            my_arr(i, j) = (i + 1) * (1 + j)
        Next j
    Next i
    
    'exemple how to use sub
    Call set_mtx_on_cell(my_arr, Application.Caller.Address, ActiveWorkbook.Name, ActiveSheet.Name)

End Function




