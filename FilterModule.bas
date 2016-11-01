Attribute VB_Name = "FilterModule"
Public Sub quick_selection_filter(ictrl As IRibbonControl)
    quick_selection_filter_inner
End Sub


Private Sub quick_selection_filter_inner()
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    
    With fh
    
        .filtruj_po_selekcji
    
    End With

    Set fh = Nothing
End Sub


Public Sub quick_clear(ictrl As IRibbonControl)
    inner_quick_clear
End Sub

Public Sub inner_quick_clear()
    
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    fh.ustaw_swiezy_filtr
    
    Set fh = Nothing
End Sub
