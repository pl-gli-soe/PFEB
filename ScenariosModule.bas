Attribute VB_Name = "ScenariosModule"
' set config for osea
Public Sub scenario1(ictrl As IRibbonControl)
    'MsgBox "scenario1"
    
    ' =====================================================
    
    'ms7p5100    DLY_COM 4
    'ms7p5100    PLAN_COM    5
    'ms7p5100    DLY_XMIT    6
    'ms7p5100    PLANNING_XMIT   7
    'ms7p5100    FORECAST_FORMAT 8
    
    'ms7p5200    STD_PACK    9
    'ms7p5200    SCH_PACK    10
    'ms7p5200    RECV_TYPE   11
    
    'ms7p5900    MODE    12
    'ms7p5900    TRANSIT 13
    'ms7p5900    TC  14
    'ms9pv400    STD_PACK    15
    
    'zk7ptogl    CHECK_FRST_CURRENT_TYPE 16
    'zk7ptogl    CHECK_FUTURE_CURRENT_TYPE   17
    
    MsgBox "scenario 1: OSEA column order"
    set_this_scenario 1
    refresh_startup_form
    ' =====================================================

End Sub


' show all possible labels in report
Public Sub scenario2(ictrl As IRibbonControl)
    'MsgBox "scenario2"
    ' =====================================================
    
    'ms7p3100    'DEPT    '4
    'ms7p3100    'OPER    '5
    'ms7p3100    'PART_NAME   6
    'ms7p3100    'SCHED_PUBLISHED 7
    'ms7p3100    'SCHED_POINT 8
    'ms7p5100    'DLY_COM 9
    'ms7p5100    'PLAN_COM    '10
    'ms7p5100    'SUPP_NM 11
    'ms7p5100    'COUNTRY_CD  12
    'ms7p5100    'ZIP 13
    'ms7p5100    'DLY_XMIT    '14
    'ms7p5100    'PLANNING_XMIT   15
    'ms7p5100    'FORECAST_FORMAT 16
    'ms7p5100    'ALIAS   17
    'ms7p5100    'ABRV_NM 18
    'ms7p5200    'DESC    '19
    'ms7p5200    'fieldNAME   20
    'ms7p5200    'STD_PACK    '21
    'ms7p5200    'SCH_PACK    '22
    'ms7p5200    'RECV_TYPE   23
    'ms7p5200    'FUT_RECV_TYPE   24
    'ms7p5200    'PAY_TYPE    '25
    'ms7p5200    'BEGIN_DATE  26
    'ms7p5200    'PSUP_TYPE_P_S   27
    'ms7p5900    'SUPP_NM 28
    'ms7p5900    'NM  29
    'ms7p5900    'PRIM    '30
    'ms7p5900    'SEC 31
    'ms7p5900    'MODE    '32
    'ms7p5900    'TRANSIT 33
    'ms7p5900    'TC  34
    'ms9pv400    'STD_PACK    '35
    'ms9pv400    'CNTR    '36
    'ms9pv400    'GR_CONT_WG  37
    'ms9pv400    'UNIT_WT 38
    'mysprh40    'ROUTE1  39
    'mysprh40    'DOCK1   40
    'ms9pop00    'BANK    '41
    'ms9pop00    'A   42
    'ms9pop00    'F_U 43
    'ms9pop00    'DK  44
    'ms9pop00    'PCS_TO_GO   45
    'zk7ptogl    'CHECK_FRST_CURRENT_TYPE 46
    'zk7ptogl    'CHECK_FUTURE_CURRENT_TYPE   47
    'zk7ptogl    'CHECK_FRST_MUL  48
    'mysptog0    CHECK_FRST_CURRENT_TYPE   49
    'mysptog0    CHECK_FUTURE_CURRENT_TYPE 50
    'mysptog0    CHECK_FRST_MUL 51



    MsgBox "scenario 2: all columns"
    set_this_scenario 2
    refresh_startup_form
    ' =====================================================
End Sub

Public Sub scenario3(ictrl As IRibbonControl)
    'MsgBox "scenario3"
    
    MsgBox "scenario 3: FMA"
    set_this_scenario 3
    refresh_startup_form
End Sub

Public Sub scenario4(ictrl As IRibbonControl)
    'MsgBox "scenario4"
    
    MsgBox "scenario 4: Component"
    set_this_scenario 4
    refresh_startup_form
End Sub

Public Sub scenario5(ictrl As IRibbonControl)
    'MsgBox "scenario5"
    
    MsgBox "scenario 5: BTN scenario"
    set_this_scenario 5
    refresh_startup_form
End Sub

Public Sub scenario6(ictrl As IRibbonControl)
    'MsgBox "scenario6"
    MsgBox "scenario x: custom column order"
    set_this_scenario 6
    refresh_startup_form
End Sub

Public Sub scenario7(ictrl As IRibbonControl)
    'MsgBox "scenario7"
    MsgBox "scenario x: custom column order"
    set_this_scenario 7
    refresh_startup_form
End Sub

Public Sub scenario8(ictrl As IRibbonControl)
    'MsgBox "scenario8"
    MsgBox "scenario x: custom column order"
    set_this_scenario 8
    refresh_startup_form
End Sub

Public Sub scenario9(ictrl As IRibbonControl)
    'MsgBox "scenario8"
    MsgBox "scenario x: custom column order"
    set_this_scenario 9
    refresh_startup_form
End Sub


Public Sub set_this_scenario(s As Integer)
    

    
    Dim cs As Worksheet, rs As Worksheet
    
    Set cs = ThisWorkbook.Sheets("config")
    Set rs = ThisWorkbook.Sheets("register")
    
    Dim cR As Range, rr As Range, r As Range
    Set cR = cs.Range("A2").End(xlDown)
    Set cR = cs.Range(cs.Range("A2"), cR)
    Set cR = cR.Offset(0, 2)
    
    Set rr = rs.Range("A2").End(xlDown)
    Set rr = rs.Range(rs.Range("A2"), rr).Offset(0, s)
    
    If rr.COUNT = cR.COUNT Then
    
        i = 1
        Do
            cR.item(i) = rr.item(i)
            i = i + 1
        Loop While i < (rr.COUNT + 1)
        
        
        StartupForm.set_saved_priv_variable True
        StartupForm.LabelStatus.Caption = "Status: changes saved!"
    Else
        ' a jesli nie moze sie pokazac nie powinienem obslugiwac zadnej logiki
        ' StartupForm.set_saved_priv_variable False
        MsgBox "ten msgbox nigdy nie powinien sie pokazac"
    End If
End Sub

Public Sub cfg(ictrl As IRibbonControl)
    open_startup_form
End Sub



' this sub is mainly for changing visible columns in output report
Public Sub open_startup_form()

    refresh_startup_form
    StartupForm.show vbModeless

End Sub

Public Sub refresh_startup_form()
    With StartupForm
        .ListBox1.Clear
        .ListBox2.Clear
    
        Dim cs As Worksheet, cR As Range, r As Range
        
        Set cs = ThisWorkbook.Sheets("config")
        Set cR = cs.Range("A2").End(xlDown)
        Set cR = cs.Range(cs.Range("A2"), cR)
        Set cR = cR.Offset(0, 2)
        
        For Each r In cR
            If r = "" Then
                .ListBox2.AddItem r.Offset(0, 1)
            End If
        Next r
        
        x = PIERWSZY_MOZLIWY_NUMER_DO_SETU
        Do
        
            For Each r In cR
                If CStr(r.Value) = CStr(x) Then
                    .ListBox1.AddItem r.Offset(0, 1)
                    
                    Exit For
                End If
            Next r
            
            x = x + 1
            
        Loop Until x > cR.COUNT
        
        .Repaint
    End With
End Sub
