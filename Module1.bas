Attribute VB_Name = "Module1"
Sub markred()
Attribute markred.VB_ProcData.VB_Invoke_Func = " \n14"
'
' markred Macro
'

'
    Range("G5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("V2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "OSEA"
    Range("U2").Select
    Selection.Copy
    Range("V2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
