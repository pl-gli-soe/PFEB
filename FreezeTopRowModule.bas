Attribute VB_Name = "FreezeTopRowModule"
Public Sub freeze_top_row()
Attribute freeze_top_row.VB_ProcData.VB_Invoke_Func = " \n14"
'

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub
