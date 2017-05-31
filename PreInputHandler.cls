VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreInputHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Private mh As MgoHandler
Private preish As Worksheet, ish As Worksheet, r As Range, ir As Range

Private fh As FilterHandling


Private contracts As Dictionary
Private cont As ContractItem
Private strpltpnduns As String


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
                        ir.Offset(0, 1) = mh.m_zk7pcont.pn
                        ir.Offset(0, 2) = mh.m_zk7pcont.iDUNS(Int(x))
                        ir.Offset(0, 3) = mh.m_zk7pcont.iA(Int(x))
                        ir.Offset(0, 4) = mh.m_zk7pcont.iPO_NO(Int(x))
                        ir.Offset(0, 5) = mh.m_zk7pcont.iDATE_ADDED(Int(x))
                        ir.Offset(0, 6) = mh.m_mgo.convertToDateFromMS9POP00Date(mh.m_zk7pcont.iEFF(Int(x)))
                        ir.Offset(0, 7) = mh.m_mgo.convertToDateFromMS9POP00Date(mh.m_zk7pcont.iEXP(Int(x)))
                        
                        ' OBSOLETE LOGIC
                        ' sprawdz_czy CDate(ir.Offset(0, 6)), CDate(ir.Offset(0, 7)), ir.Offset(0, 8)
                        ir.Offset(0, 9) = mh.m_zk7pcont.iSUPPLIER_NAME(Int(x))
                        
                        
                        If IsNumeric(Left(mh.m_zk7pcont.iBYR(Int(x)), 1)) Then
                        
                            ir.Offset(0, 10) = """'" & mh.m_zk7pcont.iBYR(Int(x))
                        Else
                            ir.Offset(0, 10) = mh.m_zk7pcont.iBYR(Int(x))
                        End If
                        
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
    
    
    
    
    ' -----------------------------------------------------------------------------------------------------
End Sub


Public Sub contract_decriptions()
    
    ' z racji nowego podjescia do zagdnienia kontraktow
    ' z poprzedniej petli (powyzej) schowana zostala logika przypisania komentarza do okresu
    ' trwania danego iterowanego kontraktu
    ' okazuje sie ze logika wymaga aby istniala mozliwosc konfronotwania informacji pomiedzy wierszami!
    
    ' OBIETKOWE PODEJSCIE DO PROBLEMU
    ' -----------------------------------------------------------------------------------------------------
    
    ' 0. stworzenie slownika w ktorym tymczasowo beda sie znajdowac ulozone informacji na temat kazdego
    ' kontraktu ktory zostal zrzucony do akursza input
    Set contracts = New Dictionary
    ' 1. zaciagniecie informacji na temat wybranego kontraktu do slownika w petli arkusza input
    Set ir = ish.Range("a2")
    ir.Select
    Do
        
        ' ITERACJA DLA KOLEJNEGO KONTRAKTU
        ' -------------------------------------------------
        
        'strpltpnduns = CStr(ir) & CStr(ir.Offset(0, 1)) ' & CStr(ir.Offset(0, 2))
        ' jednak tylko strPLTPN - bo duns przeciez moze sie zmienic
        ' ir.Select
        strpltpn = CStr(ir) & CStr(ir.Offset(0, 1)) ' & CStr(ir.Offset(0, 2))
        Set cont = Nothing
        Set cont = New ContractItem
        
        Set cont.referencja = ir
        cont.plt = ir
        cont.pn = ir.Offset(0, PFEB.INPUT_PN - 1)
        cont.duns = ir.Offset(0, PFEB.INPUT_DUNS - 1)
        
        cont.eff = ir.Offset(0, PFEB.INPUT_EFF - 1)
        cont.exp = ir.Offset(0, PFEB.INPUT_EXP - 1)
        
        cont.buyer_code = ir.Offset(0, PFEB.INPUT_BYR_CODE - 1)
        
        ' logika dodawania elementu do slownika
        If contracts.Exists(CStr(strpltpn)) Then
            contracts(CStr(strpltpn)).add_contract cont
        Else
            
            contracts.Add strpltpn, New PnForContractItem
            contracts(CStr(strpltpn)).add_contract cont
        End If
        
        
        Set cont = Nothing
        ' -------------------------------------------------
        Set ir = ir.Offset(1, 0)
    Loop Until Trim(ir) = ""
    
    
    Dim pncs As PnForContractItem
    
    For Each klucz In contracts.Keys
        Set pncs = contracts(klucz)
        pncs.is_this_pn_have_contract_with_P
        pncs.create_status_desc_dict
        pncs.primary_check_dates_logic
        pncs.secondary_check_successors_for_curr_active_project
        pncs.third_define_overcoming_param
        pncs.fourth_check_successors_with_dunses
        pncs.put_finally_description_on_ref
        
    Next
End Sub

Private Sub Class_Initialize()
    Set preish = ThisWorkbook.Sheets("preinput")
    Set ish = ThisWorkbook.Sheets("input")
    Set r = preish.Range("b2")
    Set ir = ish.Range("a2")
    
    
    Set fh = New FilterHandling
    
    Set pns_and_contracts = New Dictionary
    
End Sub

Private Sub Class_Terminate()
    Set fh = Nothing
    Set pns_and_contracts = Nothing
End Sub


' to jest obsoletowa funkcjonalnosc ale zostawie dla wzoru dla przyszlego algorytmu
Private Sub sprawdz_czy(eff As Date, exp As Date, ByRef input_r As Range)

    If exp < Date And Trim(CStr(input_r.Offset(0, -5))) = "P" Then
        input_r = "CONTACT BUYER NOW!"
        input_r.Interior.Color = RGB(255, 0, 0)
        
    ElseIf exp < Date Then
        input_r = "EXPIRED"
        input_r.Interior.Color = RGB(255, 79, 79)
    ElseIf Year(exp) = Year(Date) Then
        input_r = "EXPIRES THIS YEAR"
        input_r.Interior.Color = RGB(238, 253, 47)
        
        If eff > Date Then
            input_r = "EXPIRES THIS YEAR, BUT EFF IN FUTURE"
            input_r.Interior.Color = RGB(168, 125, 181)
        End If
        
    ElseIf exp = Date Then
        input_r = "EXPIRATION - CONTACT BUYER"
        input_r.Interior.Color = RGB(0, 255, 255)
        
    ElseIf eff > Date Then
        input_r = "FUTURE OK"
        input_r.Interior.Color = RGB(153, 243, 253)
        
    ElseIf eff = Date Then
        input_r = "CONTRACT STARTS TODAY"
        input_r.Interior.Color = RGB(57, 223, 245)
        
    ElseIf eff < Date And exp > Date Then
       input_r = "OK - VALID"
       input_r.Interior.Color = RGB(65, 255, 61)
    End If
    
End Sub

Public Sub dostosuj_layout_preinput()
    With ish
        '.Columns("I:J").ColumnWidth = 9.29
        '.Columns("I:J").Font.size = 8
        
        
        .Columns("I:I").WrapText = True
    End With
    
End Sub
