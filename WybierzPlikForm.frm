VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WybierzPlikForm 
   Caption         =   "Wybierz Plik typu Wizard"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4410
   OleObjectBlob   =   "WybierzPlikForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WybierzPlikForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub BtnSubmit_Click()
    inner_run
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    inner_run

End Sub


Private Sub inner_run()


        
        
        
        If Me.ListBox1.ListCount > 0 Then
            If Me.ListBox1.Value <> "" Then
                wh.nazwa_aktywnego_pliku = Me.ListBox1.Value
            Else
                MsgBox "gdzie jest Twoj W-General... koncze z Toba wspolprace"
                Exit Sub
            End If
            
        Else
            MsgBox "nie ma czego wybrac... koncze z Toba wspolprace!"
            Exit Sub
        End If
        
        
        
        If Me.ListBox1.Value <> "" Then
            
            Dim wgenh As WGeneralHandler
            Set wgenh = New WGeneralHandler
            If wgenh.przekaz_plik(Workbooks(Me.ListBox1.Value)) Then
                
                ' przekaz plik to funkcja boolowska ktora przekazuje plik i sprawdza czy to faktycznie jest w general
                ' =====================================================================================================
                wgenh.pakuj_dane_do_pre_inputu
                ' =====================================================================================================
            End If
        End If
        
        
        Set wh = Nothing
    Else
        MsgBox "gdzie jest Twoj W-General... koncze z Toba wspolprace"
        Exit Sub
    End If
End Sub

