VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PnForContractItem"
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


Private c As Collection
Private ci As ContractItem

Private pa_arr(1 To 3) As String


Private dic_status_description As Dictionary



Public Sub add_contract(mci As ContractItem)
    
    c.Add mci
End Sub

Private Sub Class_Initialize()
    Set c = New Collection
    
    '498997790 P&A Bochum
    '424217560  P&A Luton
    '498999655 P&A Russelsheim
    ' ----------------------------
    pa_arr(1) = "498997790"
    pa_arr(2) = "424217560"
    pa_arr(3) = "498999655"
    ' ----------------------------
End Sub

Private Sub Class_Terminate()
    Set c = Nothing
End Sub


Public Function is_this_pn_have_contract_with_P() As Boolean
    
    is_this_pn_have_contract_with_P = False
    
    For Each ci In c
        If CStr(Trim(ci.referencja.Offset(0, PFEB.INPUT_A - 1))) = "P" Then
            is_this_pn_have_contract_with_P = is_this_pn_have_contract_with_P Or True
            ci.is_active_in_another_words_have_p_on_a_field = True
        End If
    Next ci
End Function


Public Sub primary_check_dates_logic()
    
    
    For Each ci In c
    
        If CDate(ci.eff) < CDate(ci.exp) Then
            ' chronological contract - so fine
            ' ---------------------------------
            
            If CDate(ci.exp) < CDate(Date) Then
                ' past logic
                ' =========================================
                ci.status_description = CStr(dic_status_description("P"))
                ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(191, 191, 191)
                ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Font.Color = RGB(255, 255, 255)
                ci.is_past = True
                
                ' =========================================
            ElseIf (CDate(ci.exp) > CDate(Date)) And (CDate(ci.eff) > CDate(Date)) Then
                ' future logic
                ' =========================================
                ci.status_description = CStr(dic_status_description("F"))
                ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 255, 255)
                ci.is_future = True
                
                ' =========================================
            ElseIf (CDate(ci.exp) >= CDate(Date)) And (CDate(ci.eff) <= CDate(Date)) Then
                ' curr logic
                ' =========================================
                ci.status_description = CStr(dic_status_description("V"))
                ci.is_valid = True
                ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(102, 255, 51)
                
                ' pomimo tego ze jest valid warto sprawdzic czy nie zbliza sie exp w najblizszym czasie
                If CDate(ci.exp) <= CDate(Date) + 30 Then
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("ETM"))
                    ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(255, 0, 0)
                    ci.expires_this_month = True
                ElseIf CDate(ci.exp) <= CDate(Date) + 120 Then
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("ETQ"))
                    ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(255, 255, 102)
                    ci.expires_this_quarter = True
                ElseIf Year(CDate(ci.exp)) = Year(CDate(Date)) Then
                    ' expires this year
                    ci.expires_this_year = True
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("ETY"))
                    ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(255, 255, 0)
                    
                    ' expires in the next year or later
                    ' ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("EL"))
                End If
                
                
                
                
                ' =========================================
            Else
                ci.status_description = CStr(dic_status_description("UNDEF"))
            End If
            
            ' ---------------------------------
        ElseIf CDate(ci.eff) = CDate(ci.exp) Then
            
            ' eff equal exp - kontrakt jednodniowy - co z tym robimy?
            
            If CDate(ci.eff) = CDate(Date) Then
                ci.status_description = CStr(dic_status_description("TODAY IS THE DAY"))
                ci.is_valid = True
            Else
                If CDate(ci.eff) < CDate(Date) Then
                    ci.status_description = CStr(dic_status_description("ONE DAY IN PAST"))
                    ci.is_past = True
                ElseIf CDate(ci.eff) > CDate(Date) Then
                    ci.status_description = CStr(dic_status_description("ONE DAY IN FUTURE"))
                    ci.is_future = True
                End If
            End If
            
            
        ' this situation is cleary not fine
        ' when: eff > exp - it's  not chronological
        Else
            ci.status_description = CStr(dic_status_description("NOT CHRONO"))
        End If
        
    Next ci
End Sub

Public Sub secondary_check_successors_for_curr_active_project()


    Dim data_skonczenia_validu As Date
    Dim tmp_active_ci As ContractItem
    
    For Each ci In c
        
        If ci.is_active_in_another_words_have_p_on_a_field And ci.is_valid Then
            
            Set tmp_active_ci = ci
            data_skonczenia_validu = ci.exp
            Exit For
        End If
    Next ci
    
    For Each ci In c
        If (ci.is_future Or ci.is_valid) And (Not ci.is_active_in_another_words_have_p_on_a_field) Then
            
            
            ' 1. rowne daty poczatku i konca
            If CDate(data_skonczenia_validu) = ci.eff Then
            
                ci.is_successor = True
                tmp_active_ci.has_future_successor = True
                
                ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("IS"))
                ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(156, 235, 254)
                
                tmp_active_ci.status_description = _
                    tmp_active_ci.status_description & " AND " & CStr(dic_status_description("WS"))
                    
               If tmp_active_ci.expires_this_year Then
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 204, 153)
                    
                ElseIf tmp_active_ci.expires_this_month Then
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 80, 0)
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Font.Color = RGB(255, 255, 255)
                ElseIf tmp_active_ci.expires_this_quarter Then
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 180, 0)
                End If
                
            End If
            
            ' 2. jeden dzien roznicy
            If CDate(data_skonczenia_validu) + 1 = ci.eff Then
                
                ci.is_successor = True
                tmp_active_ci.has_future_successor = True
                
                ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("IS"))
                ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(156, 235, 254)
                
                
                tmp_active_ci.status_description = _
                    tmp_active_ci.status_description & " AND " & CStr(dic_status_description("WS"))
                
                If tmp_active_ci.expires_this_year Then
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 204, 153)
                    
                ElseIf tmp_active_ci.expires_this_month Then
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 80, 0)
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Font.Color = RGB(255, 255, 255)
                ElseIf tmp_active_ci.expires_this_quarter Then
                    tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 180, 0)
                
                End If
            End If
        End If
    Next ci
End Sub


Public Sub third_define_overcoming_param()


    Dim prev_ci As ContractItem
    
    Set prev_ci = Nothing
    
    For Each ci In c
    
        If Not prev_ci Is Nothing Then
            
            diff = ci.eff - prev_ci.exp
            prev_ci.overcoming = diff
            
            prev_ci.referencja.Offset(0, PFEB.INPUT_OVERCOMING - 1).Value = diff
            
            If diff <> 0 Then
                prev_ci.referencja.Offset(0, PFEB.INPUT_OVERCOMING - 1).Interior.Color = RGB(240, 240, 240)
            End If
        End If
            
        Set prev_ci = ci
    Next ci
    
    
End Sub

Public Sub fourth_check_successors_with_dunses()


    Dim data_skonczenia_validu As Date
    Dim tmp_active_ci As ContractItem
    
    For Each ci In c
        
        If ci.is_active_in_another_words_have_p_on_a_field And ci.is_valid Then
            
            Set tmp_active_ci = ci
            Exit For
        End If
    Next ci
    
    
    If Not tmp_active_ci Is Nothing Then
    
        For Each ci In c
        
            ' sprawdzmy osobno P&A
            If sprawdz_najpierw_P_and_A(CStr(ci.duns)) Then
            
                ' ci to jest sprawdzany contract item
                ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("PA"))
                'ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 80, 0)
                'ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Font.Color = RGB(255, 255, 255)
            End If
            
            
        
            If ci.is_successor Then
                ' mamy juz sukcesora nie ma potrzeby dodatkowego wpisu
                
            ElseIf Not ci.is_successor Then
        
                If (ci.is_future Or ci.is_valid) And (Not ci.is_active_in_another_words_have_p_on_a_field) Then
                    
                    ci.is_potential_successor = True
                    tmp_active_ci.has_potential_future_successor = True
                    

                    If CStr(ci.duns) = CStr(tmp_active_ci.duns) Then
                    
                    
                        
                    
                        ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("PS"))
                        tmp_active_ci.status_description = _
                            tmp_active_ci.status_description & " AND " & CStr(dic_status_description("WPS"))
                            
                        'tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 80, 0)
                        'tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Font.Color = RGB(255, 255, 255)
    
                    ElseIf CStr(ci.duns) <> CStr(tmp_active_ci.duns) Then
                    
                        ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("PSDD"))
                        tmp_active_ci.status_description = _
                            tmp_active_ci.status_description & " AND " & CStr(dic_status_description("WPSDD"))
                            
                        'tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Interior.Color = RGB(0, 80, 0)
                        'tmp_active_ci.referencja.Offset(0, PFEB.INPUT_DESC - 1).Font.Color = RGB(255, 255, 255)
                    End If
                End If
            End If
        Next ci
    
    End If

End Sub

Private Function sprawdz_najpierw_P_and_A(duns As String) As Boolean
    
    ' pesymistycznie z gory neguj
    sprawdz_najpierw_P_and_A = False
    
    
    ' przejrzyj_tabele_dunsow_z_p_and_a
    For x = LBound(pa_arr) To UBound(pa_arr)
        If CStr(duns) = CStr(pa_arr(x)) Then
            sprawdz_najpierw_P_and_A = True
        End If
    Next x
End Function

Public Sub put_finally_description_on_ref()
    
    For Each ci In c
        ci.referencja.Offset(0, PFEB.INPUT_DESC - 1) = CStr(ci.status_description)
    Next ci
End Sub

Public Sub create_status_desc_dict()
    Set dic_status_description = New Dictionary
    
    dic_status_description.Add "TODAY IS THE DAY", "TODAY IS THE DAY"
    dic_status_description.Add "ONE DAY IN PAST", "ONE DAY LONG CONTRACT FROM PAST"
    dic_status_description.Add "ONE DAY IN FUTURE", "ONE DAY LONG CONTRACT FROM FUTURE"
    dic_status_description.Add "NOT CHRONO", "NOK: NOT CHRONOLOGICAL EFF Date is bigger then EXP Date"
    dic_status_description.Add "UNDEF", "NOT DEFINED LOGIC"
    dic_status_description.Add "V", "VALID"
    dic_status_description.Add "P", "EXPIRED"
    dic_status_description.Add "F", "FUTURE"
    dic_status_description.Add "ETM", "EXPIRES WITHIN 30 DAYS"
    dic_status_description.Add "ETQ", "EXPIRES THIS QUARTER"
    dic_status_description.Add "ETY", "EXPIRES THIS YEAR"
    dic_status_description.Add "ETYORL", "EXPIRES IN THE NEXT YEAR OR LATER"
    dic_status_description.Add "WS", "WITH SUCCESSOR"
    dic_status_description.Add "WPS", "WITH POTENCIAL SUCCESSOR"
    dic_status_description.Add "WPSDD", "WITH POTENCIAL SUCCESSOR BUT WITH DIFFERENT DUNS"
    dic_status_description.Add "WOS", "WITHOUT SUCCESSOR"
    dic_status_description.Add "IS", "IS SUCCESSOR"
    dic_status_description.Add "PS", "IS POTENCIAL SUCCESSOR"
    dic_status_description.Add "PSDD", "IS POTENCIAL SUCCESSOR BUT WITH DIFFERENT DUNS"
    dic_status_description.Add "PA", "P&A DUNS MATCH"
    
    
End Sub

Public Sub clear_status_desc_dict()
    dic_status_description.RemoveAll
    Set dic_status_description = Nothing
End Sub
