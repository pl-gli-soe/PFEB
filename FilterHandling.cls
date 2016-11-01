VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fup_code As String

Private fup_code_ok As Boolean

Dim r As Range, ir As Range

Public fma_resp_cfg As Range
Public my_fup_code_cfg As Range
Public selection_cfg As Range


Public Sub ustaw_swiezy_filtr()

    If ActiveSheet.AutoFilterMode = True Then
    Else
        Range("A1").AutoFilter
    End If
    On Error Resume Next
    Range("A1").Parent.ShowAllData
End Sub


Public Function filtruj_po_selekcji() As Boolean

    If Selection.COUNT = 1 Then
        filtruj_po_selekcji = filtruj_po_aktywnej_komorce()
    Else
        filtruj_po_selekcji = filtruj_po_aktywnych_komorkach()
    End If

End Function


Private Function filtruj_po_aktywnych_komorkach() As Boolean
    filtruj_po_aktywnych_komorkach = False
    
    Dim tmp As Range
    Dim v(1 To SELECTION_LIMIT) As String
    Set tmp = Selection
        
    
        If tmp.Columns.COUNT = 1 Then
        
            Dim ir As Range
            i = 1
            For Each ir In tmp
                If Not ir.EntireRow.Hidden Then
                    v(i) = ir
                    i = i + 1
                End If
            Next ir
            
        
            ' v = Application.Transpose(Selection)
            Selection.AutoFilter Field:=tmp(1, 1).Column, Criteria1:=v, Operator:=xlFilterValues
        Else
            MsgBox "nie mozesz nalozyc dwoch filtrow na raz!"
        End If
End Function

Private Function filtruj_po_aktywnej_komorce() As Boolean
    filtruj_po_aktywnej_komorce = False
    

    Range("A1").AutoFilter ActiveCell.Column, ActiveCell.Value
    filtruj_po_aktywnej_komorce = True
    
    
        
End Function

