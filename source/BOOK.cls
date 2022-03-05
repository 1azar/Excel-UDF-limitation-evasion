VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

'Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
'    MsgBox ("You have opend the " & Wb.Name)
'End Sub

Private Sub App_AfterCalculate()
    
    If (-1 = Not GLOBAL_MTX) Then
        'Debug.Print ("GLOBAL_MTX is empty")
    Else
        Dim target_range As String
        Dim init_row As Integer
        Dim init_col As Integer
        init_row = range(GLOBAL_CELL_ADDRESS).row
        init_col = range(GLOBAL_CELL_ADDRESS).Column
        'init_col = Split(range(GLOBAL_CELL_ADDRESS).Address, "$")(1)
        'init_row = Split(range(GLOBAL_CELL_ADDRESS).Address, "$")(2)
        'Debug.Print ("arr setted to global vars: " & "mtx " & UBound(GLOBAL_MTX, 1) & "x" & UBound(GLOBAL_MTX, 2) & " " & GLOBAL_CELL_ADDRESS)
        target_range = get_range_as_str(init_row, init_col, init_row + UBound(GLOBAL_MTX, 2), init_col + UBound(GLOBAL_MTX, 1))
        'Workbooks(GLOBAL_BOOK).Sheets(GLOBAL_SHEET).range(target_range).Value2 = WorksheetFunction.Transpose(GLOBAL_MTX)
        
        If (GLOBAL_LPF_evented_first <> False) Then
            GLOBAL_LPF_evented_first = False
            Workbooks(GLOBAL_BOOK).Sheets(GLOBAL_SHEET).range(target_range).Value2 = GLOBAL_MTX
        End If
    
        Erase GLOBAL_MTX
        GLOBAL_CELL_ADDRESS = ""
        GLOBAL_BOOK = ""
        GLOBAL_SHEET = ""
    End If
End Sub

Private Sub Workbook_Open()
    Set App = Application
End Sub

Function get_range_as_str(ByRef start_row As Integer, ByRef start_column As Integer, ByRef end_row As Integer, ByRef end_column As Integer) As String
    '1,2,10,20 -> "B1:T10
    Dim result_range As String
    result_range = Get_Letter_by_Number_(start_column) + CStr(start_row) + ":" + Get_Letter_by_Number_(end_column) + CStr(end_row)
    get_range_as_str = result_range
End Function

Private Function Get_Letter_by_Number_(lngCol As Integer) As String
    '1 -> "A", 2 -> "B", 4 -> "D", -3 -> "A"
    Dim vArr
    If lngCol >= 1 Then
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Get_Letter_by_Number_ = CStr(vArr(0))
    Else
        Debug.Print ("Cant define the letter for column nuber - " + CStr(lngCol) + "!")
        Get_Letter_by_Number_ = "A"
    End If
End Function
