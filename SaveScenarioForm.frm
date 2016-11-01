VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveScenarioForm 
   Caption         =   "Save Scenario"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3300
   OleObjectBlob   =   "SaveScenarioForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveScenarioForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private reg_sh As Worksheet

' logika obslugujaca zapisywania odpowiednich scenariuszy
' do arkusza register
' ----------------------------------------------------
' ----------------------------------------------------

Private Sub BtnS5_Click()

    ' in short words - we define new scenario for our purposes
    podmien_zawartosc_kolumny_ 5, BtnS5, CheckBox5

End Sub

Private Sub BtnS6_Click()

    ' in short words - we define new scenario for our purposes
    podmien_zawartosc_kolumny_ 6, BtnS6, CheckBox6

End Sub

Private Sub BtnS7_Click()

    ' in short words - we define new scenario for our purposes
    podmien_zawartosc_kolumny_ 7, BtnS7, CheckBox7

End Sub

Private Sub BtnS8_Click()

    ' in short words - we define new scenario for our purposes
    podmien_zawartosc_kolumny_ 8, BtnS8, CheckBox8

End Sub

Private Sub BtnS9_Click()

    ' in short words - we define new scenario for our purposes
    podmien_zawartosc_kolumny_ 9, BtnS9, CheckBox9

End Sub


' args:
' podmien_zawartosc_kolumny_
' kk - ktora kolumna
Private Sub podmien_zawartosc_kolumny_(kk As Integer, btn As Variant, cB As Variant)
    
    Dim cs As Worksheet, rs As Worksheet
    
    Set cs = ThisWorkbook.Sheets("config")
    Set rs = ThisWorkbook.Sheets("register")
    
    Dim cR As Range, rr As Range, r As Range
    Set cR = cs.Range("A2").End(xlDown)
    Set cR = cs.Range(cs.Range("A2"), cR)
    Set cR = cR.Offset(0, 2)
    
    Set rr = rs.Range("A2").End(xlDown)
    Set rr = rs.Range(rs.Range("A2"), rr).Offset(0, kk)
    
    If rr.COUNT = cR.COUNT Then
    
        i = 1
        Do
            rr.item(i) = cR.item(i)
            i = i + 1
        Loop While i < (rr.COUNT + 1)
        
        
        StartupForm.set_saved_priv_variable True
        StartupForm.LabelStatus.Caption = "Status: changes saved!"
        
        btn.Enabled = False
        cB.Value = False
        btn.Caption = "Saved on " & CStr(kk) & "th slot"
        
    Else
        ' a jesli nie moze sie pokazac nie powinienem obslugiwac zadnej logiki
        ' StartupForm.set_saved_priv_variable False
        MsgBox "ten msgbox nigdy nie powinien sie pokazac"
    End If
End Sub

' koniec logiki obslugujacej
' zapisywania scenariuszy do arkuszu register
' ----------------------------------------------------
' ----------------------------------------------------

' logika obslugujaca checkboxy
' ----------------------------------------------------

Private Sub CheckBox5_Click()
    If CheckBox5.Value Then
        BtnS5.Enabled = True
        BtnS5.Caption = "O'write this slot"
    Else
        BtnS5.Enabled = False
        BtnS5.Caption = "Save on 5rd slot"
    End If
End Sub

Private Sub CheckBox6_Click()
    If CheckBox6.Value Then
        BtnS6.Enabled = True
        BtnS6.Caption = "O'write this slot"
    Else
        BtnS6.Enabled = False
        BtnS6.Caption = "Save on 6rd slot"
    End If
End Sub

Private Sub CheckBox7_Click()
    If CheckBox7.Value Then
        BtnS7.Enabled = True
        BtnS7.Caption = "O'write this slot"
    Else
        BtnS7.Enabled = False
        BtnS7.Caption = "Save on 7rd slot"
    End If
End Sub

Private Sub CheckBox8_Click()
    If CheckBox8.Value Then
        BtnS8.Enabled = True
        BtnS8.Caption = "O'write this slot"
    Else
        BtnS8.Enabled = False
        BtnS8.Caption = "Save on 8rd slot"
    End If
End Sub

Private Sub CheckBox9_Click()
    If CheckBox9.Value Then
        BtnS9.Enabled = True
        BtnS9.Caption = "O'write this slot"
    Else
        BtnS9.Enabled = False
        BtnS9.Caption = "Save on 9rd slot"
    End If
End Sub

' koniec logiki obslugujacej checkboxy
' ----------------------------------------------------



' konstruktor - poczatek
' ----------------------------------------------------
Private Sub UserForm_Initialize()
    Set reg_sh = ThisWorkbook.Sheets("register")
    
    Dim tmp As Range ' , item As Control
    Set tmp = reg_sh.Range("a2")
    Set tmp = reg_sh.Range(tmp, tmp.End(xlDown))
    
    ' for all 9 scenarios
    For x = 5 To 9
    
        ' inside
        ' now will check all scenarios how content looks like
        If Application.WorksheetFunction.COUNT(tmp.Offset(0, Int(x))) > 0 Then
            ' it means that this scenario have some content
            For Each item In Me.Controls
                If item.NAME Like "*BtnS*" & CStr(x) & "*" Then
                    item.Enabled = False
                    Exit For
                End If
            Next item
        Else
            
        End If
    Next x
End Sub

' konstruktor - koniec
' ----------------------------------------------------
