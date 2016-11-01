VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartupForm 
   Caption         =   "Init"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   OleObjectBlob   =   "StartupForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private saved As Boolean
Private sth_changed As Boolean

Public Sub set_saved_priv_variable(s As Boolean)
    saved = s
End Sub

Public Function show_saved_priv_variable()
    show_saved_priv_variable = saved
End Function

Private Sub Btn_5th_scenario_Click()
    set_this_scenario 5
    refresh_startup_form
End Sub

Private Sub Btn_6th_scenario_Click()
    set_this_scenario 6
    refresh_startup_form
End Sub

Private Sub Btn_7th_scenario_Click()
    set_this_scenario 7
    refresh_startup_form
End Sub

Private Sub Btn_8th_scenario_Click()
    set_this_scenario 8
    refresh_startup_form
End Sub

Private Sub Btn_9th_scenario_Click()
    set_this_scenario 9
    refresh_startup_form
End Sub

Private Sub Btn_Component_scenario_Click()
    ' MsgBox "scenariusz nie zostal jeszcze zdefiniowany"
    set_this_scenario 4
    refresh_startup_form
End Sub

Private Sub Btn_FMA_scenario_Click()

    ' MsgBox "scenariusz nie zostal jeszcze zdefiniowany"
    set_this_scenario 3
    refresh_startup_form
End Sub

Private Sub BtnAllColumns_Click()
    set_this_scenario 2
    refresh_startup_form
End Sub

Private Sub BtnOsea_Click()
    set_this_scenario 1
    refresh_startup_form
End Sub

Private Sub BtnAllLeft_Click()

    
    
    With StartupForm
        .ListBox1.Clear
        .ListBox2.Clear
    
        Dim cs As Worksheet, cR As Range, r As Range
        
        Set cs = ThisWorkbook.Sheets("config")
        Set cR = cs.Range("A2").End(xlDown)
        Set cR = cs.Range(cs.Range("A2"), cR)
        Set cR = cR.Offset(0, 2)
        
        For Each r In cR
            .ListBox1.AddItem r.Offset(0, 1)
        Next r
        
        .Repaint
    End With
    
    saved = False
    sth_changed = True
    StartupForm.LabelStatus.Caption = "Status: changes unsaved!"
End Sub

Private Sub BtnAllRight_Click()

    

    With StartupForm
        .ListBox1.Clear
        .ListBox2.Clear
    
        Dim cs As Worksheet, cR As Range, r As Range
        
        Set cs = ThisWorkbook.Sheets("config")
        Set cR = cs.Range("A2").End(xlDown)
        Set cR = cs.Range(cs.Range("A2"), cR)
        Set cR = cR.Offset(0, 2)
        
        For Each r In cR
            .ListBox2.AddItem r.Offset(0, 1)
        Next r
        
        .Repaint
    End With
    
    saved = False
    sth_change = True
    StartupForm.LabelStatus.Caption = "Status: changes unsaved!"
End Sub

Private Sub BtnCancel_Click()
    Me.hide
    If saved Then
    ElseIf (Not saved) And sth_changed Then
        
        answer = MsgBox("Close without saving?", vbYesNo + vbQuestion, "Closing without saving")
        
        If answer = vbYes Then
            ' nothing to do!
            ' -------------------
            
            ' -------------------
            
        ElseIf answer = vbNo Then
            MsgBox "Click OK to save this configuration"
            reorganize_config_sheet
        End If
    End If
End Sub

Private Sub goleft()

    StartupForm.LabelStatus.Caption = "Status: changes unsaved!"
    saved = False
    sth_changed = True

    With Me
        For i = 0 To .ListBox2.ListCount - 1
        
            If .ListBox2.Selected(i) = True Then
                .ListBox1.AddItem CStr(.ListBox2.List(i))
                .ListBox2.RemoveItem i
            End If
        Next i
    End With
End Sub

Private Sub goright()

    StartupForm.LabelStatus.Caption = "Status: changes unsaved!"
    saved = False
    sth_changed = True

    With Me
        For i = 0 To .ListBox1.ListCount - 1
        
            If .ListBox1.Selected(i) = True Then
                .ListBox2.AddItem CStr(.ListBox1.List(i))
                .ListBox1.RemoveItem i
            End If
        Next i
    End With
End Sub

Private Sub BtnOneLeft_Click()
    goleft
End Sub

Private Sub BtnOneRight_Click()
    goright
End Sub



Private Sub BtnSaveOnly_Click()
    hide
    saved = True
    sth_changed = False
    StartupForm.LabelStatus.Caption = "Status: changes saved!"
    reorganize_config_sheet
End Sub

Private Sub BtnSaveScenario_Click()

    ' this action will be assigned with with openinig
    ' new form to perform final saving on proper button
    ' =============================================================
    saved = True
    sth_changed = False
    StartupForm.LabelStatus.Caption = "Status: changes saved!"
    reorganize_config_sheet
    SaveScenarioForm.show
    ' =============================================================
End Sub

Private Sub BtnSubmit_Click()
    hide
    saved = True
    sth_changed = False
    StartupForm.LabelStatus.Caption = "Status: changes saved!"
    reorganize_config_sheet
    inner_main
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    goright
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    goleft
End Sub

Private Sub UserForm_Initialize()
    saved = False
    sth_changed = False
    StartupForm.LabelStatus.Caption = "Status: no changes"
    
    ' -------------------------------------------------------------------
    ' check scenarios
    ' need to think if i really want it
    
    Set reg_sh = ThisWorkbook.Sheets("register")
    
    Dim tmp As Range ' , item As Control
    Set tmp = reg_sh.Range("a2")
    Set tmp = reg_sh.Range(tmp, tmp.End(xlDown))
    
    ' for all 9 scenarios
    For x = 3 To 9
    
        ' inside
        ' now will check all scenarios how content looks like
        
            ' it means that this scenario have some content
            For Each item In Me.Controls
            
                ' Btn_7th_scenario
                If item.NAME Like "*Btn_*" & CStr(x) & "*_scenario" Then
                    If Application.WorksheetFunction.COUNT(tmp.Offset(0, Int(x))) > 0 Then
                        
                        'item.Enabled = True
                        'item.BackColor = &H8000000F
                    ElseIf Application.WorksheetFunction.COUNT(tmp.Offset(0, Int(x))) = 0 Then
                        'item.Enabled = True
                        'item.BackColor = &H80000012
                    End If
                End If
            Next item
        
    Next x
    
    
    '
    '
    ' -------------------------------------------------------------------
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    hide
    If saved Then
    ElseIf Not saved Then
        
        answer = MsgBox("Close without saving?", vbYesNo + vbQuestion, "Closing without saving")
        
        If answer = vbYes Then
            ' nothing to do!
            ' -------------------
            
            ' -------------------
            
        ElseIf answer = vbNo Then
            MsgBox "Click OK to save this configuration"
            reorganize_config_sheet
        End If
    End If

End Sub

Private Sub reorganize_config_sheet()
    With StartupForm
    
        Dim cs As Worksheet, cR As Range, r As Range
        Set cs = ThisWorkbook.Sheets("config")
        Set cR = cs.Range("A2").End(xlDown)
        Set cR = cs.Range(cs.Range("A2"), cR)
        Set cR = cR.Offset(0, 2)
        
        For Each r In cR
            r = ""
        Next r

        For i = 0 To .ListBox1.ListCount - 1
            
            tmp = .ListBox1.List(i)
            For Each r In cR
                If Trim(r.Offset(0, 1)) = Trim(tmp) Then
                    r = i + 4
                    Exit For
                End If
            Next r
        Next i
    End With
End Sub
