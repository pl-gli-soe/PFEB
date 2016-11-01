Attribute VB_Name = "GlobalModule"
Global Const SELECTION_LIMIT = 256
Global Const PIERWSZY_MOZLIWY_NUMER_DO_SETU = 4
Global Const BLANK = "BLANK"
Global Const SZEROKOSC_KOLUMNY = 17
Global Const PN_LEN = 8
Global Const DUNS_LEN = 9

Global Const NAZWA_ARKUSZA_STARTOWEGO = "STEP-BY-STEP"



Public Function add_zeros(arg, ile) As String
    add_zeros = ""
    
    If Len(arg) = ile Then
        add_zeros = CStr(arg)
    ElseIf Len(arg) < ile Then
    
        Dim tmp As String
        tmp = Application.WorksheetFunction.Rept("0", ile - Len(arg))
        
        add_zeros = CStr(tmp) & CStr(arg)
    Else
        MsgBox "sth wrong with the length of the string in sub add_zeros"
        add_zeros = ""
    End If
End Function
