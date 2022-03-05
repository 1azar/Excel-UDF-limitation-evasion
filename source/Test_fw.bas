Attribute VB_Name = "Test_fw"
Public Function generate_matrix(h As Integer, w As Integer)
    Dim my_arr() As Variant
    ReDim my_arr(0 To w, 0 To h)
    For i = 0 To UBound(my_arr, 1)
        For j = 0 To UBound(my_arr, 2)
            my_arr(i, j) = CStr(j) & " " & CStr(i)
        Next j
    Next i
    
    'exemple how to use the sub
    Call set_mtx_on_cell(my_arr, Application.Caller.Address, ActiveWorkbook.Name, ActiveSheet.Name)

End Function
