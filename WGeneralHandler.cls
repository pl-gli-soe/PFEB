VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WGeneralHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private plik As Workbook


Public Function przekaz_plik(p As Workbook) As Boolean

    przekaz_plik = False
    ' najpiew sprawdz czy w ogole ma sens co przekazywac
    
    ' sprawdze 4 pierwsze labelki - powinno wystarczyc
    ' w general pattern
    If Trim(p.ActiveSheet.Cells(1, 1)) = "F-up Code" Then
        If Trim(p.ActiveSheet.Cells(1, 2)) = "Plant Code" Then
            If Trim(p.ActiveSheet.Cells(1, 3)) = "Part Number" Then
                If Trim(p.ActiveSheet.Cells(1, 4)) = "Part Name" Then
            
                    Set plik = p
                    przekaz_plik = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    Set plik = Nothing
End Function

Private Sub Class_Initialize()
    Set plik = Nothing
End Sub

Private Sub Class_Terminate()
    Set plik = Nothing
End Sub


Public Sub pakuj_dane_do_pre_inputu()
    
    Dim r As Range, pish As Worksheet
    Set r = plik.ActiveSheet.Cells(1, 1)
    Set pish = ThisWorkbook.Sheets(PFEB.NAZWA_ARKUSZA_PRE_INPUT)
    
    Do
        ' body kolejnych wierszy w generala
        ' ---------------------------------------------
        
        ' ---------------------------------------------
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub
