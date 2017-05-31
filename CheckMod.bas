Attribute VB_Name = "CheckMod"
Public Sub CheckOsea()

    ' nasz kod adjustment kolorystyczny bledu
    
    to_jest_dla_kolumny_STD_PACK
    'to_jest_dla_kolumny_FLAG_CURR
    'to_jest_dla_kolumny_FLAG_FUT
    'to_jest_dla_kolumny_DLY_COM
    'to_jest_dla_kolumny_PLAN_COM
    'to_jest_dla_kolumny_RECV_T
    to_jest_dla_kolumny_DLY_XMIT
    to_jest_dla_kolumny_PLAN_XMIT
    to_jest_dla_kolumny_TMODE
    to_jest_dla_kolumny_TTIME
    to_jest_dla_kolumny_FOR_FOR
    to_jest_dla_kolumny_STCODE
    
    ' ij specjalne warunki "CGH"
    ' ============================
     'mega_uniwersalna_petla 8, "CGH", 255, 235, 156, E_CGH_LOGIC
     'mega_uniwersalna_petla 8, "CGH TSK", 255, 235, 156, E_CGH_LOGIC
     'mega_uniwersalna_petla 8, "TSK", 255, 235, 156, E_CGH_LOGIC
     'mega_uniwersalna_petla 8, "TSK CGH", 255, 235, 156, E_CGH_LOGIC
     
     cztery_petle_zlozenia 8, "CGH", 9, 255, 235, 156
     cztery_petle_zlozenia 8, "CGH TSK", 9, 255, 235, 156
     cztery_petle_zlozenia 8, "TSK", 9, 255, 235, 156
     cztery_petle_zlozenia 8, "TSK CGH", 9, 255, 235, 156
     mega_uniwersalna_petla 8, "", 255, 235, 156, E_EQUAL
     mega_uniwersalna_petla 9, "", 255, 235, 156, E_EQUAL
     
    ' ============================
    ' specjalny warunek dla TCODE GME + ZA
    mega_uniwersalna_petla 15, "", 255, 0, 0, E_ZA_OPERATOR
    ' ============================
    ' specjalny warunek dla KB
    ' mega_uniwersalna_petla 6, "M", 255, 0, 0, E_FLAG_OPERATOR
    mega_uniwersalna_petla 0, "M", 255, 0, 0, E_FLAG_OPERATOR_OSEA
    mega_uniwersalna_petla 0, "", 255, 0, 0, E_FLAG_OPERATOR_OSEA_F
    
    petla_zlozenia_pol 6, "W", 7, "M", 169, 208, 142
    petla_zlozenia_pol 0, "KB", 6, "M", 255, 255, 255
    
    ' ============================
    ' specjalny warunek dla KB RECV TYPE
    mega_uniwersalna_petla 10, "", 255, 0, 0, E_RECV_OPERATOR
    
    ' pod component
    kb_loop
    
    ' com codes implementation
    com_code_loop
    
    wsadz_w_kolumne_vlookup
    
   
    PFEB.dTT_loop
    
    'kolor if no FU
    mega_uniwersalna_petla 23, "no FU", 211, 27, 93, E_EQUAL

    
    ' Font.Color = RGB(255, 0, 0)
End Sub

Public Sub kb_loop()
    
    'prawdzenie component schedule published
    
    Dim r As Range
    Set r = Cells(3, 21)
    Do
        If Trim(r.Parent.Cells(r.Row, 1)) = "KB" Then
            If Trim(r) = "Y" Then
            Else
                r.Interior.Color = RGB(255, 0, 0)
            End If
            
        Else
            If Trim(r) = "Y" Then
                r.Interior.Color = RGB(255, 0, 0)
            Else
                
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r.Parent.Cells(r.Row, 1)) = ""
End Sub

Public Sub com_code_loop()
    
    Dim r As Range
    Set r = Cells(3, 1)
    
    kolumna_dly_com_code = 9
    kolumna_plan_com_code = 10
    kolumna_country_code = 20
    
    Do
    
        porownaj_com_comdy r, kolumna_dly_com_code, kolumna_plan_com_code, kolumna_country_code
        iteracja_com_code r, kolumna_dly_com_code, kolumna_country_code
        iteracja_com_code r, kolumna_plan_com_code, kolumna_country_code
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
End Sub

Public Sub dTT_loop()
    Dim r As Range
    Set r = Cells(3, 1)
    
    kolumna_TT = 15
    Do
        porownaj_TTme r, kolumna_TT
        
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub

Private Sub porownaj_TTme(r As Range, kolumna)
    Dim dTT_sh As Worksheet
    Set dTT_sh = ThisWorkbook.Sheets("DEFAULT TT")
    
    Dim dTT_r As Range
    Set dTT_r = dTT_sh.Range("a2")
    
    Do
        ' sprawdzenie czy planty sa takie same
        If Trim(dTT_r) = Trim(r) Then
            If Trim(dTT_r.Offset(0, 2)) = Trim(r.Offset(0, 2)) Then
            
                With r.Offset(0, kolumna - 1)
                    .Interior.Color = RGB(255, 235, 156)
                    .Font.Bold = True
                    .Font.Color = RGB(10, 10, 230)
                End With
                
            End If
        
        End If
        
        
        Set dTT_r = dTT_r.Offset(1, 0)
    Loop Until Trim(dTT_r) = ""
    
End Sub



Private Sub iteracja_com_code(r As Range, kolumna, kolumna_country_code)


    com_code = Trim(r.Parent.Cells(r.Row, kolumna))
    
    ' Debug.Assert com_code <> "DIJ"
    
    this_cc = Trim(r.Parent.Cells(r.Row, kolumna_country_code))
    
    
    'to jest dla com code TSK
    If this_cc = "KR" Or this_cc = "CN" Or this_cc = "HK" Then
    
      
        If PFEB.RegularExprModule.match_with_tsk(com_code) = E_NO_MATCH Then
            
            ' to jest NOK
            ' w ogole nie istnieje zaden wpis typu TSK
             r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
            
        
        ElseIf PFEB.RegularExprModule.match_with_tsk(com_code) >= E_TSK_3_PRE And _
            PFEB.RegularExprModule.match_with_tsk(com_code) <= E_TSK_2_END Then
            
            
            ' to jest OK
            e = PFEB.RegularExprModule.match_with_tsk(com_code)
            
            Select Case e
                Case PFEB.E_TSK_3_PRE ' OK
                    ' logika pod "TSK xxx xxx"
                    
                Case PFEB.E_TSK_3_MID ' OK
                    ' "xxx TSK xxx"
                    
                Case PFEB.E_TSK_3_END ' OK
                    ' "xxx xxx TSK"
                    
                Case PFEB.E_TSK_2_PRE ' OK
                    ' "TSK xxx"
                    
                Case PFEB.E_TSK_2_END ' OK
                    ' "xxx TSK"
                    
                    
            End Select
            
        Else
        
            If PFEB.RegularExprModule.match_with_tsk(com_code) = E_TSK_ALONE Then
                ' to jest NOK => "TSK"
                ' ale mozna juz okreslic ze to jest specyficzny NOK
                r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
                
            ElseIf PFEB.RegularExprModule.match_with_tsk(com_code) > 300 Then
                
                ' nie wiem co to: TSK pojawilo sie 3 razy
                ' TSK TSK TSK
                r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
                
            ElseIf PFEB.RegularExprModule.match_with_tsk(com_code) > 200 Then
            
                ' nie wiem co to: TSK pojawilo sie dwa razy
                ' np: TSK TSK xxx
                ' np: TSK xxx TSK
                ' np: xxx TSK TSK
                ' lub
                ' np: TSK TSK
                r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    
        
        
    'to jest dla com code CGH
    ElseIf this_cc = "VN" Or this_cc = "MY" Or this_cc = "JP" Or this_cc = "AU" Or this_cc = "IN" Or this_cc = "US" Or this_cc = "CA" Or this_cc = "MX" Then
    
        
         If PFEB.RegularExprModule.match_with_cgh(com_code) = E_NO_MATCH Then
            
            ' to jest NOK
            ' w ogole nie istnieje zaden wpis typu CGH
            r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
            
        
        ElseIf PFEB.RegularExprModule.match_with_cgh(com_code) >= E_CGH_3_PRE And _
            PFEB.RegularExprModule.match_with_cgh(com_code) <= E_CGH_2_END Then
            
            
            ' to jest OK
            e = PFEB.RegularExprModule.match_with_cgh(com_code)
            
            Select Case e
                Case PFEB.E_CGH_3_PRE ' OK
                    ' logika pod "CGH xxx xxx"
                    
                Case PFEB.E_CGH_3_MID ' OK
                    ' "xxx CGH xxx"
                    
                Case PFEB.E_CGH_3_END ' OK
                    ' "xxx xxx CGH"
                    
                Case PFEB.E_CGH_2_PRE ' OK
                    ' "CGH xxx"
                    
                Case PFEB.E_CGH_2_END ' OK
                    ' "xxx CGH"
                    
                    
            End Select
            
        Else
        
            If PFEB.RegularExprModule.match_with_cgh(com_code) = E_CGH_ALONE Then
                ' to jest NOK => "CGH"
                ' ale mozna juz okreslic ze to jest specyficzny NOK
                r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
            ElseIf PFEB.RegularExprModule.match_with_cgh(com_code) > 3000 Then
                
                ' nie wiem co to: CGH pojawilo sie 3 razy
                ' CGH CGH CGH
                r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
            ElseIf PFEB.RegularExprModule.match_with_cgh(com_code) > 2000 Then
            
                ' nie wiem co to: TSK pojawilo sie dwa razy
                ' np: CGH CGH xxx
                ' np: CGH xxx CGH
                ' np: xxx CGH CGH
                ' lub
                ' np: CGH CGH
                r.Offset(0, kolumna - 1).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    Else
    
        ' jeszcze cos innego ...
        ' -----------------------------------------------------------
        ''
      
        
        ' -----------------------------------------------------------
    End If

End Sub

Private Sub porownaj_com_comdy(r As Range, kolumna_dly, kolumna_plan, kolumna_cc)


    ' r as range - jako pierwszy argument tego suba odnosi sie do pierwszej kolumny raportu (zaczynamy od 3 wiersza danych)

    Dim r_dly As Range, r_plan As Range, r_cc As Range
    
    
    ' komorka z dly kolumny
    Set r_dly = r.Parent.Cells(r.Row, kolumna_dly)
    ' komorka z com code plan kolumny
    Set r_plan = r.Parent.Cells(r.Row, kolumna_plan)
    ' kormorka z country code kolumny
    Set r_cc = r.Parent.Cells(r.Row, kolumna_cc)
    
    
    dly = Trim(r_dly.Value)
    plan = Trim(r_plan.Value)
    
    If Trim(dly) = Trim(plan) Then
        ' sa identyczne wiec spoko - mozemy np nalozyc zielony font
        With r_dly.Font
            .Italic = False
            .Color = RGB(0, 0, 0)
        End With
        
        With r_plan.Font
            .Italic = False
            .Color = RGB(0, 0, 0)
        End With
    Else
    
        ' na dzien dobry okazuje sie ze te dwa com code sa od siebie rozne - mozemy nalozyc pierwszy layer kolorow
        ' nalezy pamietac ze po tej instrukcji sa kolejne ktore moga zakryc postawiony kolor w tej czesci...
        ' --------------------------------------------------------------------------------------------------------------
        
        
        ' "TODO" here - mark r_dly (cell) and r_plan (cell) because they are different
        '----------------------------------------------------------------------
        With r_dly.Font
            .Italic = True
            .Color = RGB(10, 10, 230)
        End With
        
        With r_plan.Font
            .Italic = True
            .Color = RGB(10, 10, 230)
        End With
        ' --------------------------------------------------------------------------------------------------------------
    End If
End Sub

Public Sub CheckFMA()

    ' nasz kod adjustment kolorystyczny bledu
    
    ' to jest dla kolumny E STD PACK FMA
    mega_uniwersalna_petla 4, 0, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 4, 1, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 4, "", 255, 93, 93, E_EQUAL
    
    
    ' to jest dla kolumny H TTIME FMA
    mega_uniwersalna_petla 7, "", 255, 93, 93, E_EQUAL

    ' to jest dla kolumny I TTCOUNT FMA
    ' mega_uniwersalna_petla 8, "5", 255, 0, 0, E_NOT_EQUAL
    mega_uniwersalna_petla 8, "", 255, 93, 93, E_ZA_OPERATOR
    
    ' to jest dla kolumny K FLAG_CURR and COMP
    mega_uniwersalna_petla 10, "M", 255, 93, 93, E_FLAG_OPERATOR
    
    ' recv
    ' to jest dla kolumny G RECV_T FMA
    ' mega_uniwersalna_petla 6, "", 255, 0, 0, E_NOT_EQUAL
    mega_uniwersalna_petla 6, "", 255, 93, 93, E_RECV_OPERATOR
       
    ' to jest dla kolumny L FLAG_FUT FMA
    mega_uniwersalna_petla 11, "", 255, 93, 93, E_NOT_EQUAL
    
     ' to jest dla kolumny N C_FLAG_FUT FMA
    mega_uniwersalna_petla 13, "", 255, 93, 93, E_NOT_EQUAL
    
     ' to jest dla kolumny O DLY_COM FMA
    'mega_uniwersalna_petla 14, "", 255, 0, 0, E_EQUAL
    'mega_uniwersalna_petla 14, "CGH", 255, 0, 0, E_EQUAL
    
    ' to jest dla kolumny P PLAN_COM FMA
    'mega_uniwersalna_petla 15, "", 255, 0, 0, E_EQUAL
    'mega_uniwersalna_petla 15, "CGH", 255, 0, 0, E_EQUAL
    
    ' to jest dla kolumny S BANK FMA
    mega_uniwersalna_petla 18, "", 255, 93, 93, E_EQUAL
   
   ' to jest dla kolumny T ROUTE FMA
    mega_uniwersalna_petla 19, "", 255, 93, 93, E_EQUAL
    
    
    ' ij specjalne warunki "CGH"
    ' ============================
    'mega_uniwersalna_petla 14, "CGH", 255, 0, 0, E_CGH_LOGIC
    'mega_uniwersalna_petla 14, "CGH TSK", 255, 235, 156, E_CGH_LOGIC
    'mega_uniwersalna_petla 14, "TSK", 255, 235, 156, E_CGH_LOGIC
    'mega_uniwersalna_petla 14, "TSK CGH", 255, 235, 156, E_CGH_LOGIC
    cztery_petle_zlozenia 14, "CGH", 15, 255, 235, 156
    cztery_petle_zlozenia 14, "CGH TSK", 15, 255, 235, 156
    cztery_petle_zlozenia 14, "TSK", 15, 255, 235, 156
    cztery_petle_zlozenia 14, "TSK CGH", 15, 255, 235, 156
    ' ============================
End Sub
Public Sub CheckCOMP()

    ' wykorzystanie petli dla zwyklej flagi ze wzgledu na podobna zasade dzialania reguly
    to_jest_dla_kolumny_FLAG_CURR
    to_jest_dla_kolumny_FLAG_FUT
    
    component_rules
     'ij specjalne warunki "CGH"
    ' ============================
    mega_uniwersalna_petla 9, "CGH", 255, 0, 0, E_CGH_LOGIC
     cztery_petle_zlozenia 9, "CGH", 10, 255, 0, 0
     cztery_petle_zlozenia 9, "CGH TSK", 10, 255, 0, 0
     cztery_petle_zlozenia 9, "TSK", 10, 255, 0, 0
     cztery_petle_zlozenia 9, "TSK CGH", 10, 255, 0, 0
     mega_uniwersalna_petla 9, "", 255, 0, 0, E_EQUAL
     mega_uniwersalna_petla 9, "", 255, 0, 0, E_EQUAL
     mega_uniwersalna_petla 10, "", 255, 0, 0, E_EQUAL
     mega_uniwersalna_petla 10, "", 255, 0, 0, E_EQUAL
    ' ============================
End Sub

Public Sub CheckBTN()

    mega_uniwersalna_petla 5, 0, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 5, 1, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 5, "", 255, 93, 93, E_EQUAL

    mega_uniwersalna_petla 4, 0, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 4, "", 255, 93, 93, E_EQUAL
    
    mega_uniwersalna_petla 6, 0, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 6, "_________", 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 6, "", 255, 93, 93, E_EQUAL
    
    mega_uniwersalna_petla 7, 0, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 7, "________", 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 7, "", 255, 93, 93, E_EQUAL
    
    mega_uniwersalna_petla 8, 0, 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 8, "__________", 255, 93, 93, E_EQUAL
    mega_uniwersalna_petla 8, "", 255, 93, 93, E_EQUAL
End Sub

Private Sub mega_uniwersalna_petla(o_ile_w_prawo, txt_do_sprawdzenia, cR, cG, cB, jaki_operator As E_JAKI_OPERTOR)

    Dim r As Range
    Set r = Range("a3")
    Do
        If jaki_operator = E_EQUAL Then
            If Trim(r.Offset(0, o_ile_w_prawo)) = Trim(txt_do_sprawdzenia) Then
                r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
                
                If CStr(txt_do_sprawdzenia) = "no FU" Then
                    r.Offset(0, o_ile_w_prawo).Font.Color = RGB(255, 255, 255)
                Else
                    r.Offset(0, o_ile_w_prawo).Font.Color = RGB(0, 0, 0)
                End If
                
            End If
        ElseIf jaki_operator = E_NOT_EQUAL Then
            If Trim(r.Offset(0, o_ile_w_prawo)) <> Trim(txt_do_sprawdzenia) Then
                r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
            End If
        ElseIf jaki_operator = E_LIKE Then
            If CStr(txt_do_sprawdzenia) Like "*" & CStr(r.Offset(0, o_ile_w_prawo)) & "*" Then
                r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
            End If
        ElseIf jaki_operator = E_NOT_LIKE Then
            If Not (CStr(txt_do_sprawdzenia) Like "*" & CStr(r.Offset(0, o_ile_w_prawo)) & "*") Then
                r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
            End If
        ElseIf jaki_operator = E_ZA_OPERATOR Then
            
            ' ta czesc kodu odpowiedzialna '
            ' jest za odpowiednie rozroznienie plt ZA z perspektywy TTIME'u
            ' nalezy pamietac ze TC = 7 dla ZA, for rest TC = 5
            
            If r = "ZA" Then
                If r.Offset(0, o_ile_w_prawo) <> 7 Then
                    r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
                End If
            Else
                If r.Offset(0, o_ile_w_prawo) <> 5 Then
                    r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
                End If
            End If
            
        ElseIf jaki_operator = E_CGH_LOGIC Then
        
            If Trim(r.Offset(0, o_ile_w_prawo)) = CStr(txt_do_sprawdzenia) Then
                If Trim(r.Offset(0, o_ile_w_prawo + 1)) = "" Or Trim(r.Offset(0, o_ile_w_prawo + 1)) = CStr(txt_do_sprawdzenia) Then
            
                    Range(r.Offset(0, o_ile_w_prawo), r.Offset(0, o_ile_w_prawo + 1)).Interior.Color = RGB(cR, cG, cB)
                
                End If
            End If
            
        ElseIf jaki_operator = E_FLAG_OPERATOR Then
        
            If Trim(r) = "KB" Then
                
                ' tutaj uruchamiamy logike jesli plt jest componentowy
                ' ======================================================
                ''
                '
                If Trim(r.Offset(0, o_ile_w_prawo + 2)) <> CStr(txt_do_sprawdzenia) Then
                    r.Offset(0, o_ile_w_prawo + 2).Interior.Color = RGB(cR, cG, cB)
                End If
                '
                ''
                ' ======================================================
            Else
                If Trim(r.Offset(0, o_ile_w_prawo)) <> CStr(txt_do_sprawdzenia) Then
                    r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
                End If
            End If
            
        ElseIf jaki_operator = E_FLAG_OPERATOR_OSEA Then
        
            If Trim(r) = "KB" Then
                
                ' tutaj uruchamiamy logike jesli plt jest componentowy w scenariuszu osea
                ' ======================================================
                ''
                '
                If Trim(r.Offset(0, 6)) <> "" Then
                    r.Offset(0, 6).Interior.Color = RGB(cR, cG, cB)
                End If
                
                ' ======================================================
            Else
                If Trim(r.Offset(0, 6)) <> CStr(txt_do_sprawdzenia) Then
                     r.Offset(0, 6).Interior.Color = RGB(cR, cG, cB)
                End If
            End If
            
            
         ElseIf jaki_operator = E_FLAG_OPERATOR_OSEA_F Then
        
            If Trim(r) = "KB" Then
                
                ' tutaj uruchamiamy logike jesli plt jest componentowy w scenariuszu osea
                ' ======================================================
                ''
                '
                If Trim(r.Offset(0, 7)) <> "" Then
                    r.Offset(0, 7).Interior.Color = RGB(cR, cG, cB)
                End If
                
                ' ======================================================
            Else
                If Trim(r.Offset(0, 7)) <> CStr(txt_do_sprawdzenia) Then
                     r.Offset(0, 7).Interior.Color = RGB(cR, cG, cB)
                End If
                
            End If
            
        ElseIf jaki_operator = E_RECV_OPERATOR Then
        
            If Trim(r) = "KB" Then
                If Trim(r.Offset(0, o_ile_w_prawo)) <> "PDC" Then
                    r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
                End If
            Else
                If Trim(r.Offset(0, o_ile_w_prawo)) <> CStr(txt_do_sprawdzenia) Then
                    r.Offset(0, o_ile_w_prawo).Interior.Color = RGB(cR, cG, cB)
                End If
            End If
        End If
           
        Set r = r.Offset(1, 0)
    Loop While Trim(r.Value) <> ""
End Sub
Private Sub cztery_petle_zlozenia(o_ile_w_prawo_1, txt_do_sprawdzenia_1, o_ile_w_prawo_2, cR, cG, cB)

     petla_zlozenia_pol o_ile_w_prawo_1, txt_do_sprawdzenia_1, 9, "CGH", cR, cG, cB
     petla_zlozenia_pol o_ile_w_prawo_1, txt_do_sprawdzenia_1, 9, "TSK CGH", cR, cG, cB
     petla_zlozenia_pol o_ile_w_prawo_1, txt_do_sprawdzenia_1, 9, "TSK", cR, cG, cB
     petla_zlozenia_pol o_ile_w_prawo_1, txt_do_sprawdzenia_1, 9, "CGH TSK", cR, cG, cB

End Sub

Private Sub petla_zlozenia_pol(o_ile_w_prawo_1, txt_do_sprawdzenia_1, o_ile_w_prawo_2, txt_do_sprawdzenia_2, cR, cG, cB)
    Dim r As Range
    Set r = Range("a3")
    Do
        If Trim(r.Offset(0, o_ile_w_prawo_1)) = CStr(txt_do_sprawdzenia_1) Then
                If Trim(r.Offset(0, o_ile_w_prawo_2)) = CStr(txt_do_sprawdzenia_2) Then
            
                     r.Offset(0, o_ile_w_prawo_1).Interior.Color = RGB(cR, cG, cB)
                     r.Offset(0, o_ile_w_prawo_2).Interior.Color = RGB(cR, cG, cB)
                End If
            End If
        Set r = r.Offset(1, 0)
    Loop While Trim(r.Value) <> ""
End Sub

Private Sub to_jest_dla_kolumny_STD_PACK()
' to jest dla kolumny E
    mega_uniwersalna_petla 4, 0, 255, 199, 206, E_EQUAL
    mega_uniwersalna_petla 4, 1, 255, 199, 206, E_EQUAL
    mega_uniwersalna_petla 4, "", 255, 199, 206, E_EQUAL
End Sub
    
Private Sub to_jest_dla_kolumny_FLAG_CURR()
    ' to jest dla kolumny G
    mega_uniwersalna_petla 6, "W", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 6, "", 255, 0, 0, E_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_FLAG_FUT()
    ' to jest dla kolumny H
    mega_uniwersalna_petla 7, "", 255, 0, 0, E_NOT_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_DLY_COM()
    ' to jest dla kolumny I
    mega_uniwersalna_petla 8, "", 255, 235, 156, E_EQUAL
    mega_uniwersalna_petla 8, "CGH", 255, 235, 156, E_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_PLAN_COM()
    ' to jest dla kolumny J
    mega_uniwersalna_petla 9, "", 255, 235, 156, E_EQUAL
    mega_uniwersalna_petla 9, "CGH", 255, 235, 156, E_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_RECV_T()
     ' to jest dla kolumny K
    mega_uniwersalna_petla 10, "", 255, 0, 0, E_NOT_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_DLY_XMIT()
    ' to jest dla kolumny L
    mega_uniwersalna_petla 11, "Y", 255, 0, 0, E_NOT_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_PLAN_XMIT()
    ' to jest dla kolumny M
    mega_uniwersalna_petla 12, "Y", 255, 0, 0, E_NOT_EQUAL
End Sub
 
Private Sub to_jest_dla_kolumny_TMODE()
    ' to jest dla kolumny N
    mega_uniwersalna_petla 13, "AEVE", 255, 0, 0, E_NOT_LIKE
    mega_uniwersalna_petla 13, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 13, "A", 255, 0, 0, E_EQUAL
    
End Sub

Private Sub to_jest_dla_kolumny_TTIME()
    ' to jest dla kolumny O
    mega_uniwersalna_petla 14, "", 255, 0, 0, E_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_TCOUNT()
    ' to jest dla kolumny P
    mega_uniwersalna_petla 15, "5", 255, 0, 0, E_NOT_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_FOR_FOR()
    ' to jest dla kolumny Q
    mega_uniwersalna_petla 16, "M", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 16, "", 255, 0, 0, E_EQUAL
End Sub

Private Sub to_jest_dla_kolumny_STCODE()
    ' to jest dla kolumny R
    mega_uniwersalna_petla 17, "", 255, 0, 0, E_EQUAL
End Sub

Private Sub component_rules()
    'sprawdzenie bledow dla scenariusza component
    mega_uniwersalna_petla 4, 0, 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 4, 1, 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 4, "", 255, 0, 0, E_EQUAL
    'comp sched published
    mega_uniwersalna_petla 8, "Y", 255, 0, 0, E_NOT_EQUAL
    ' recv type rule -PDC
    mega_uniwersalna_petla 11, "PDC", 255, 0, 0, E_NOT_EQUAL
    ' component DLY_XMIT
    mega_uniwersalna_petla 12, "Y", 255, 0, 0, E_NOT_EQUAL
    ' component PLT_XMIT
    mega_uniwersalna_petla 13, "Y", 255, 0, 0, E_NOT_EQUAL
    'component com codes
    mega_uniwersalna_petla 9, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 10, "CGH", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 9, "CGH", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 10, "", 255, 0, 0, E_EQUAL
    'T MODE
    mega_uniwersalna_petla 14, "AEVE", 255, 0, 0, E_NOT_LIKE
    mega_uniwersalna_petla 14, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 14, "A", 255, 0, 0, E_EQUAL
    'T Time
    mega_uniwersalna_petla 15, "", 255, 0, 0, E_EQUAL
    ' T count
    mega_uniwersalna_petla 16, "5", 255, 0, 0, E_NOT_EQUAL
    ' forecast format
    mega_uniwersalna_petla 17, "W", 255, 0, 0, E_NOT_EQUAL
    'ship time code
    mega_uniwersalna_petla 18, "", 255, 0, 0, E_EQUAL
    
End Sub


Public Sub CheckALL()

    'check supplier Alias & abr name
    mega_uniwersalna_petla 5, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 6, "", 255, 0, 0, E_EQUAL

    'sprawdzenie bledow dla scenariusza All, std packs
    mega_uniwersalna_petla 7, 0, 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 7, 1, 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 7, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 9, 0, 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 9, 1, 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 9, "", 255, 0, 0, E_EQUAL
    
    ' component DLY_XMIT
    mega_uniwersalna_petla 10, "Y", 255, 0, 0, E_NOT_EQUAL
    ' component PLT_XMIT
    mega_uniwersalna_petla 11, "Y", 255, 0, 0, E_NOT_EQUAL

    ' forecast format
    mega_uniwersalna_petla 15, "W", 255, 0, 0, E_NOT_EQUAL
    'ship time code
    mega_uniwersalna_petla 16, "", 255, 0, 0, E_EQUAL
    
    ' com codes
    mega_uniwersalna_petla 17, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 18, "CGH", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 17, "CGH", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 18, "", 255, 0, 0, E_EQUAL
    'T MODE
    mega_uniwersalna_petla 19, "AEVE", 255, 0, 0, E_NOT_LIKE
    mega_uniwersalna_petla 19, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 19, "A", 255, 0, 0, E_EQUAL
    'T Time
    mega_uniwersalna_petla 20, "", 255, 0, 0, E_EQUAL
    
    ' recv
    ' to jest dla  RECV_T all
    ' mega_uniwersalna_petla 6, "", 255, 0, 0, E_NOT_EQUAL
    mega_uniwersalna_petla 22, "", 255, 0, 0, E_RECV_OPERATOR
    mega_uniwersalna_petla 23, "", 255, 0, 0, E_RECV_OPERATOR
    
    ' to jest dla kolumny  FLAG_FUT all
    mega_uniwersalna_petla 26, "", 255, 0, 0, E_NOT_EQUAL
    ' to jest dla kolumny  FLAG_ all
    mega_uniwersalna_petla 25, "", 255, 0, 0, E_NOT_EQUAL
        ' to jest dla kolumny  C_FLAG_FUT all
    mega_uniwersalna_petla 28, "", 255, 0, 0, E_NOT_EQUAL
    ' to jest dla kolumny  C_FLAG_ all
    mega_uniwersalna_petla 29, "", 255, 0, 0, E_NOT_EQUAL
    
    
    ' to jest dla kolumny  BANK
    mega_uniwersalna_petla 31, "", 255, 0, 0, E_EQUAL
    mega_uniwersalna_petla 31, 0, 255, 0, 0, E_EQUAL
   
   ' to jest dla kolumny  ROUTE
    mega_uniwersalna_petla 39, "", 255, 0, 0, E_EQUAL
    
    
    ' ij specjalne warunki "CGH"
    ' ============================
    mega_uniwersalna_petla 17, "CGH", 255, 0, 0, E_CGH_LOGIC
    
    
    't code ZA
    mega_uniwersalna_petla 21, "", 255, 0, 0, E_ZA_OPERATOR
    ' ============================
    ' specjalny warunek dla KB
    ' mega_uniwersalna_petla 6, "M", 255, 0, 0, E_FLAG_OPERATOR
    'mega_uniwersalna_petla 0, "M", 255, 0, 0, E_FLAG_OPERATOR_OSEA
    'mega_uniwersalna_petla 0, "", 255, 0, 0, E_FLAG_OPERATOR_OSEA_F
    ' ============================
    
    ' specjalny warunek dla KB RECV TYPE
    mega_uniwersalna_petla 22, "", 255, 0, 0, E_RECV_OPERATOR
    
    'KB czy plant komponentowy w all scenario
    mega_uniwersalna_petla 13, "Y", 255, 0, 0, E_RECV_OPERATOR
    
    
            'If Trim(r) = "KB" Then
                
                ' tutaj uruchamiamy logike jesli plt jest componentowy w scenariuszu all
                ' ======================================================
                ''
                '
                'If Trim(r.Offset(0, 13)) <> "Y" Then
                    'r.Offset(0, 13).Interior.Color = RGB(cR, cG, cB)
                'End If
            'End If
    
End Sub

Public Sub wsadz_w_kolumne_vlookup()
    
    Dim szablon_vlookupa As Range
    Set szablon_vlookupa = ThisWorkbook.Sheets("VLOOKUP_handler").Range("VLOOKUP")
    
    Dim general As Range
    Set general = ThisWorkbook.Sheets("DUNS").Range("A1:B8000")
    
    
    Dim r As Range
    Set r = ThisWorkbook.ActiveSheet.Range("C3")
    
    Do
        If r.Offset(0, -2) <> "ZA" Then
        
            tmp = szablon_vlookupa.Formula
            
            tmp = CStr(Replace(tmp, "-1", Replace(r.AddressLocal, "$", "")))
            tmp = CStr(Replace(tmp, "-2", general.Parent.NAME & "!" & general.Address))
            tmp = CStr(Replace(tmp, "-3", 2))
            ' w przypadku modyfikacji scenariusza osea (dodawanie / usuwanie kolumn) nalezy pamietac o odpowiednim zadjustowaniu offset
            '(po dodaniu kolumny comp sched publ konieczne bylo przesuniecie z 20 na 21)
            r.Offset(0, 21).Formula = tmp
            
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub
