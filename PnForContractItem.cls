VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PnForContractItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private c As Collection
Private ci As ContractItem


Private dic_status_description As Dictionary



Public Sub add_contract(mci As ContractItem)
    
    c.Add mci
End Sub

Private Sub Class_Initialize()
    Set c = New Collection
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
                ci.is_past = True
                
                ' =========================================
            ElseIf (CDate(ci.exp) > CDate(Date)) And (CDate(ci.eff) > CDate(Date)) Then
                ' future logic
                ' =========================================
                ci.status_description = CStr(dic_status_description("F"))
                ci.is_future = True
                
                ' =========================================
            ElseIf (CDate(ci.exp) > CDate(Date)) And (CDate(ci.eff) < CDate(Date)) Then
                ' curr logic
                ' =========================================
                ci.status_description = CStr(dic_status_description("V"))
                ci.is_valid = True
                
                ' pomimo tego ze jest valid warto sprawdzic czy nie zbliza sie exp w najblizszym czasie
                If CDate(ci.exp) < CDate(Date) + 30 Then
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("ETM"))
                ElseIf CDate(ci.exp) < CDate(Date) + 120 Then
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("ETQ"))
                ElseIf Year(CDate(ci.exp)) = Year(CDate(Date)) Then
                    ' expires this year
                    ci.expires_this_year = True
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("ETY"))
                Else
                    
                    ' expires in the next year or later
                    ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("EL"))
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
        If ci.is_future And (Not ci.is_active_in_another_words_have_p_on_a_field) Then
            
            
            ' 1. rowne daty poczatku i konca
            If CDate(data_skonczenia_validu) = ci.eff Then
            
                ci.is_successor = True
                tmp_active_ci.has_future_successor = True
                
                ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("IS"))
                tmp_active_ci.status_description = _
                    tmp_active_ci.status_description & " AND " & CStr(dic_status_description("WS"))
                
            End If
            
            ' 2. jeden dzien roznicy
            If CDate(data_skonczenia_validu) + 1 = ci.eff Then
                
                ci.is_successor = True
                tmp_active_ci.has_future_successor = True
                
                ci.status_description = ci.status_description & " AND " & CStr(dic_status_description("IS"))
                tmp_active_ci.status_description = _
                    tmp_active_ci.status_description & " AND " & CStr(dic_status_description("WS"))
            End If
        End If
    Next ci
End Sub


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
    dic_status_description.Add "ETM", "EXPIRES THIS MONTH"
    dic_status_description.Add "ETQ", "EXPIRES THIS QUARTER"
    dic_status_description.Add "ETY", "EXPIRES THIS YEAR"
    dic_status_description.Add "ETYORL", "EXPIRES IN THE NEXT YEAR OR LATER"
    dic_status_description.Add "WS", "WITH SUCCESSOR"
    dic_status_description.Add "WOS", "WITHOUT SUCCESSOR"
    dic_status_description.Add "IS", "IS SUCCESSOR"
End Sub

Public Sub clear_status_desc_dict()
    dic_status_description.RemoveAll
    Set dic_status_description = Nothing
End Sub
