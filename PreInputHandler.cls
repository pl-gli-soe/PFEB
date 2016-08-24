VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreInputHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mh As MgoHandler
Private preish As Worksheet, ish As Worksheet, r As Range, ir As Range

Private fh As FilterHandling


Public Sub clear_input_sheet()

    ' clear filtru wymaga aktywnosci danego arkusza

    Set ish = ThisWorkbook.Sheets("input")
    ish.Activate
    fh.ustaw_swiezy_filtr
    
    Set ir = ish.Range("b2").End(xlDown)
    Set ir = ir.Parent.Range("a2:z" & ir.Row)
    ir.Clear
    
End Sub

Public Sub start_(ByRef mmh As MgoHandler)
    
    
    Set mh = mmh
    
    ' first row
    ish.Activate
    Set ir = ish.Range("a2")
    Set r = preish.Range("b2")
    
    Do
    
        If Not r.EntireRow.Hidden Then
    
            ' iteracja
            ' ------------------------------
            mh.m_zk7pcont.open_this_screen
            mh.m_zk7pcont.sPN CStr(r)
            If Trim(CStr(r.Offset(0, -1))) <> "" Then mh.m_zk7pcont.sPLT CStr(r.Offset(0, -1))
            If Trim(CStr(r.Offset(0, 1))) <> "" Then mh.m_zk7pcont.sDUNS CStr(r.Offset(0, 1))
            mh.m_zk7pcont.submit
            
            ' petla w petli
            ' przegladamy teraz dane ekranu lecimy po 4 i po F8 po tym
            ' ================================================================
            Do
                ' petla petla petla
                ' 4 po kolei
                For x = 0 To 3
                    
                    If Trim(mh.m_zk7pcont.iPLT(Int(x))) <> "" Then
                        ir = mh.m_zk7pcont.iPLT(Int(x))
                        ir.Offset(0, 1) = mh.m_zk7pcont.PN
                        ir.Offset(0, 2) = mh.m_zk7pcont.iDUNS(Int(x))
                        ir.Offset(0, 3) = mh.m_zk7pcont.iA(Int(x))
                        ir.Offset(0, 4) = mh.m_zk7pcont.iPO_NO(Int(x))
                        ir.Offset(0, 5) = mh.m_zk7pcont.iDATE_ADDED(Int(x))
                        ir.Offset(0, 6) = mh.m_mgo.convertToDateFromMS9POP00Date(mh.m_zk7pcont.iEFF(Int(x)))
                        ir.Offset(0, 7) = mh.m_mgo.convertToDateFromMS9POP00Date(mh.m_zk7pcont.iEXP(Int(x)))
                        sprawdz_czy CDate(ir.Offset(0, 6)), CDate(ir.Offset(0, 7)), ir.Offset(0, 8)
                        ir.Offset(0, 9) = mh.m_zk7pcont.iSUPPLIER_NAME(Int(x))
                        ir.Offset(0, 10) = mh.m_zk7pcont.iBYR(Int(x))
                        
                        Set ir = ir.Offset(1, 0)
                        ir.Select
                        
                    Else
                        Exit For
                    End If
                    
                Next x
           
        Loop While mh.m_zk7pcont.go_down()
        ' ================================================================
        
        ' ------------------------------
        
         End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub

Private Sub Class_Initialize()
    Set preish = ThisWorkbook.Sheets("preinput")
    Set ish = ThisWorkbook.Sheets("input")
    Set r = preish.Range("b2")
    Set ir = ish.Range("a2")
    
    
    Set fh = New FilterHandling
End Sub

Private Sub Class_Terminate()
    Set fh = Nothing
End Sub

Private Sub sprawdz_czy(EFF As Date, EXP As Date, ByRef input_r As Range)

    If EXP < Date And Trim(CStr(input_r.Offset(0, -5))) = "P" Then
        input_r = "CONTACT BUYER NOW!"
        input_r.Interior.Color = RGB(255, 0, 0)
        
    ElseIf EXP < Date Then
        input_r = "EXPIRED"
        input_r.Interior.Color = RGB(255, 79, 79)
    ElseIf Year(EXP) = Year(Date) Then
        input_r = "EXPIRES THIS YEAR"
        input_r.Interior.Color = RGB(238, 253, 47)
        
        If EFF > Date Then
            input_r = "EXPIRES THIS YEAR, BUT EFF IN FUTURE"
            input_r.Interior.Color = RGB(168, 125, 181)
        End If
        
    ElseIf EXP = Date Then
        input_r = "EXPIRATION - CONTACT BUYER"
        input_r.Interior.Color = RGB(0, 255, 255)
        
    ElseIf EFF > Date Then
        input_r = "FUTURE OK"
        input_r.Interior.Color = RGB(153, 243, 253)
        
    ElseIf EFF = Date Then
        input_r = "CONTRACT STARTS TODAY"
        input_r.Interior.Color = RGB(57, 223, 245)
        
    ElseIf EFF < Date And EXP > Date Then
       input_r = "OK - VALID"
       input_r.Interior.Color = RGB(65, 255, 61)
    End If
    
End Sub

Public Sub dostosuj_layout_preinput()
    With ish
        '.Columns("I:J").ColumnWidth = 9.29
        '.Columns("I:J").Font.size = 8
        
        
        
    End With
    
End Sub
