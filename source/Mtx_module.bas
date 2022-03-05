Attribute VB_Name = "Mtx_module"
Public GLOBAL_MTX() As Variant 'indices must start from 0
Public GLOBAL_CELL_ADDRESS As String
Public GLOBAL_BOOK As String
Public GLOBAL_SHEET As String
Public GLOBAL_LPF_evented_first As Boolean


Public Sub set_mtx_on_cell(ByRef mtx() As Variant, _
                            ByRef cell As String, _
                            ByRef book_name As String, _
                            ByRef sheet_name As String)
    GLOBAL_MTX = mtx
    GLOBAL_CELL_ADDRESS = cell
    GLOBAL_BOOK = book_name
    GLOBAL_SHEET = sheet_name
    GLOBAL_LPF_evented_first = True
End Sub

Private Sub Workbook_Open()

End Sub
