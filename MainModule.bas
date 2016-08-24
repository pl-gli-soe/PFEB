Attribute VB_Name = "MainModule"
' preinput & input stuff
' =============================================================
Public Sub gen_input(ictrl As IRibbonControl)
    run_preinput_for_input
End Sub

Private Sub run_preinput_for_input()
    Dim mh As MgoHandler, pih As PreInputHandler
    Set mh = New MgoHandler
    Set pih = New PreInputHandler
    pih.clear_input_sheet
    pih.start_ mh
    pih.dostosuj_layout_preinput
    
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
