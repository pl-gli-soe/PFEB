Attribute VB_Name = "MainModule"

'
Public Sub prepare_pre_input(ictrl As IRibbonControl)
    
    With WybierzPlikForm
        
        .ListBox1.Clear
        
        For Each s In Workbooks
            .ListBox1.AddItem CStr(s.NAME)
        Next s
        
        .show
    End With
End Sub

' preinput & input stuff
' =============================================================
Public Sub gen_input(ictrl As IRibbonControl)
    run_preinput_for_input
End Sub


Public Sub test_contracts_desc()
    
    Dim pih As PreInputHandler
    Set pih = New PreInputHandler
    pih.contract_decriptions
    Set pih = Nothing
End Sub

Private Sub run_preinput_for_input()
    Dim mh As MgoHandler, pih As PreInputHandler
    Set mh = New MgoHandler
    Set pih = New PreInputHandler
    pih.clear_input_sheet
    pih.start_ mh
    pih.dostosuj_layout_preinput
    pih.contract_decriptions
    Set mh = Nothing
    
    MsgBox "ready!"
End Sub
' =============================================================


' run_main run main macro!
' =============================================================
Public Sub run_main(ictrl As IRibbonControl)
    
    inner_main
End Sub

Public Sub inner_main()
    Dim mh As MgoHandler, mrh As MainRunHandler
    Set mrh = New MainRunHandler
    Set mh = New MgoHandler
    mrh.add_new_output_sheet
    mrh.start_ mh
    
    Set mrh = Nothing
    Set mh = Nothing
    
    MsgBox "ready!"
End Sub

' =============================================================
