Attribute VB_Name = "AjdustWidthModule"
Sub osea_adjustment(ictrl As IRibbonControl)
Attribute osea_adjustment.VB_ProcData.VB_Invoke_Func = " \n14"
'
' insert_g Macro
    
    If Trim(Cells(1, 7)) = "FLAG CURRENT TYPE - M" Then
    
        ' warunek if pod dodanie nowego wiersza
        warunkowe_dodanie_wiersza
         'scalenie komorek legendy
        scalanie_pod_legende
         ' tresc legendy
        tresc_legendy
         'uklad legendy
        kreski_pod_legende
        
        ' na poczatku wszystkie kolumny maja czcionke o rozmiarze 11
        Columns("A:X").Font.size = 11
        
        'czcionka legendy
        Range("C1:U1").Font.Bold = True
        Range("C1:U1").Font.size = 14
        
        ' dodajemy czerwony nok field do legendy i srodkujemy zawartosc
        nok_field
        
        ' sekcja zmiany wielkosci kolumn oraz czcionek w tresci
        column_widths
        
        'kolorowanie  danych dla ZA na szaro
        oznaczenie_ZA
        
        ' dodanie kolumny oea
        add_osea_col
        
        'sprawdzenie bledow Osea
        CheckOsea
        
        
        
        
    Else
        MsgBox "To nie jest odpowiedni arkusz dla tego adjustmentu!"
    End If
End Sub
 

Sub fma_adjustment(ictrl As IRibbonControl)
'
' Macro1 Macro
If Trim(Cells(1, 8)) = "TRANSIT TIME" Then

    'na poczatku czcionka 11
    Columns("A:T").Font.size = 11
    
    ' szerokosc kolumn i czcionka
    FMA_columns_adj
    
    'FMA legenda
    FMA_legenda

    ' czeci ZA kolor szary
    oznaczenie_ZA

    ' sprwdzenie bledow
    CheckFMA
    
    Range("A2").Select
    
 Else
        MsgBox "To nie jest odpowiedni arkusz dla tego adjustmentu!"
 End If
 
End Sub

Sub component_adjustment(ictrl As IRibbonControl)
'
' Macro1 Macro
If Trim(Cells(1, 7)) = "COMP FLAG CURRENT TYPE - M" Then
'
    If Cells(1, 1).Value = "PLT" Then
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
    End If
    
    Columns("A:W").Font.size = 11
    
    'uklad component
    component_columns
    
    'legenda component
    component_leg
    
    'sprawdzenie component
    CheckCOMP
    
     Range("A2").Select
 Else
        MsgBox "To nie jest odpowiedni arkusz dla tego adjustmentu!"
    End If
End Sub


Sub all_adjustment(ictrl As IRibbonControl)

If Trim(Cells(1, 4)) = "PART NAME" Then
    ' dodawanie wiersza
    If Cells(1, 1).Value = "PLT" Then
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
    End If

    ' na poczatku wszystkie kolumny maja czcionke o rozmiarze 11
    Columns("A:W").Font.size = 11
    
    'all columns adjustment
    ALL_columns_width
    
    
    'kolorowanie  danych dla ZA na szaro
    Dim r As Range
    Set r = Range("a3")
    Do
        If Trim(r) = "ZA" Then
            
            ' Range("A3:W3").Interior.Color = RGB(217, 217, 217)
            Range("A" & CStr(r.Row) & ":" & "AZ" & CStr(r.Row)).Interior.Color = RGB(217, 217, 217)
                
        End If
        Set r = r.Offset(1, 0)
    Loop While Trim(r.Value) <> ""
    
    'sprawdzenie bledow
    CheckALL
    
    Range("A2").Select
   
   ' merge nok field in legend
    Range("A1:C1").Select
    Range("A1:C1").Merge
    
    ' wysrodkowanie
    With Range("A1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
   
   ' merge field in legend
    Range("D1:AZ1").Select
    Range("D1:AZ1").Merge
    
    'tresc, kolor legendy
    Cells(1, 1).Value = "NOK"
    Cells(1, 1).Interior.Color = RGB(255, 0, 0)
    Cells(1, 4).Value = "ALL COLUMNS SCENARIO"
    Cells(1, 4).Interior.Color = RGB(164, 199, 57)
    
    ' wysrodkowanie
    With Range("D1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
   'ramka legendy
   Dim rng As Range

    Set rng = Range("A1:C1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbWhite
        .Weight = xlThick
    End With
    
    'czcionka legendy
    Range("A1:D1").Font.Bold = True
    Range("A1:U1").Font.size = 16
    Cells(1, 4).Font.Color = RGB(255, 255, 255)
 Else
        MsgBox "To nie jest odpowiedni arkusz dla tego adjustmentu!"
 End If
 
End Sub

Sub BTN_adjustment(ictrl As IRibbonControl)

' insert_g Macro
    
    If Trim(Cells(1, 5)) = "UNIT WG" Then
    
        ' na poczatku wszystkie kolumny maja czcionke o rozmiarze 11
        Columns("A:W").Font.size = 11
        
        ' warunek if pod dodanie nowego wiersza
    If Cells(1, 1).Value = "PLT" Then
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
    End If
     
         'scalenie komorek legendy
        Range("A1:C1").Select
        Range("A1:C1").Merge
        Range("D1:O1").Select
        Range("D1:O1").Merge
        Cells(1, 15).Interior.Color = RGB(164, 199, 57)
        
         'uklad legendy
    Dim rng As Range
    Set rng = Range("A1:C1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
        
        'czcionka legendy
        Range("A1:U1").Font.Bold = True
        Range("A1:U1").Font.size = 14
        
        ' dodajemy czerwony nok field do legendy i srodkujemy zawartosc
        ' merge nok field in legend
        Cells(1, 1).Value = "NOK"
        Cells(1, 1).Interior.Color = RGB(255, 93, 93)
        Cells(1, 4).Value = "Packaging scenario"
        Cells(1, 4).Interior.Color = RGB(164, 199, 57)
        
        ' sekcja zmiany wielkosci kolumn oraz czcionek w tresci
        Columns("A:A").ColumnWidth = 3.71
        Columns("B:B").ColumnWidth = 8.71
        Columns("C:C").ColumnWidth = 10.29
        Columns("D:D").ColumnWidth = 32.14
        Columns("D:D").Font.size = 8
        Columns("E:E").ColumnWidth = 8.86
        Columns("F:H").ColumnWidth = 10.29
        Columns("H:H").Font.size = 8
        Columns("I:I").ColumnWidth = 11.43
        Columns("I:I").Font.size = 10
        Columns("J:J").ColumnWidth = 19
        Columns("K:K").ColumnWidth = 9.43
        Columns("L:L").ColumnWidth = 11.43
        Columns("L:L").Font.size = 10
        Columns("M:M").ColumnWidth = 13.57
        Columns("N:N").ColumnWidth = 13.86
        Columns("O:O").ColumnWidth = 19
           
        ' sprawdzanie bledow
        CheckBTN
        
        Cells(1, 1).Select
        
    Else
        MsgBox "To nie jest odpowiedni arkusz dla tego adjustmentu!"
    End If
End Sub


Private Sub column_widths()
    Columns("A:A").ColumnWidth = 3.71
    Columns("B:B").ColumnWidth = 8.71
    Columns("C:C").ColumnWidth = 10.29
    Columns("D:D").ColumnWidth = 32.14
    Columns("D:D").Font.size = 8
    Columns("E:F").ColumnWidth = 8.43
    Columns("G:H").ColumnWidth = 7.14
    Range("G2:H2").Font.size = 10
    Columns("I:J").ColumnWidth = 9.29
    Columns("I:J").Font.size = 8
    Columns("K:K").ColumnWidth = 6.29
    Columns("L:M").ColumnWidth = 7
    Columns("N:O").ColumnWidth = 10.71
    Columns("O:O").ColumnWidth = 10.71
    Columns("P:P").ColumnWidth = 7.71
    Columns("Q:Q").ColumnWidth = 6.86
    Range("Q2:Q2").Font.size = 8
    Columns("R:R").ColumnWidth = 7.29
    Columns("S:S").ColumnWidth = 9
    Columns("T:T").ColumnWidth = 7.14
    Range("T2:T2").Font.size = 9
    Range("U2:U2").Font.size = 7
    Columns("U:X").ColumnWidth = 4.86
    
End Sub

Private Sub nok_field()
    ' merge nok field in legend
    Range("U1:X1").Select
    Range("U1:X1").Merge
    
    ' wysrodkowanie
    With Range("U1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub add_osea_col()
    Range("X2").Select
    Application.CutCopyMode = False
    ActiveCell.Value = "OSEA"
    Range("U2").Select
    Selection.Copy
    Range("X2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("X2").Font.size = 11
    
    Range("A2").Select
End Sub

Private Sub kreski_pod_legende()
    Dim rng As Range

    Set rng = Range("A1:D1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
        
    Set rng = Range("E1:H1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
     
    Set rng = Range("I1:L1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
        
    Set rng = Range("M1:O1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
     End With
  
    Set rng = Range("P1:T1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
     End With
     
    Set rng = Range("U1:X1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
     End With
End Sub

Private Sub warunkowe_dodanie_wiersza()

    If Cells(1, 1).Value = "PLT" Then
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
        
        'legenda
        Cells(1, 1).Interior.Color = RGB(255, 199, 206)
        Cells(1, 5).Interior.Color = RGB(255, 232, 184)
        Cells(1, 9).Interior.Color = RGB(255, 255, 26)
        Cells(1, 13).Interior.Color = RGB(217, 217, 217)
        Cells(1, 16).Interior.Color = RGB(169, 208, 142)
        Cells(1, 21).Interior.Color = RGB(255, 0, 0)
        Cells(1, 22).Interior.Color = RGB(255, 0, 0)
        Cells(1, 23).Interior.Color = RGB(255, 0, 0)
    End If
End Sub

Private Sub scalanie_pod_legende()
    Range("A1:B1").Merge
    Range("C1:D1").Merge
    Range("F1:H1").Merge
    Range("J1:L1").Merge
    Range("N1:O1").Merge
    Range("Q1:T1").Merge
    Range("N1:O1").Merge
    Range("U1:X1").Merge
End Sub

Private Sub tresc_legendy()
    Cells(1, 3).Value = "FMA action"
    Cells(1, 6).Value = "FMA Planning action"
    Cells(1, 10).Value = "changed by O'sea"
    Cells(1, 14).Value = "ZA"
    Cells(1, 17).Value = "changed - will be OK"
    Cells(1, 21).Value = "NOK for O'sea"
End Sub

Private Sub oznaczenie_ZA()
    Dim r As Range
    Set r = Range("a3")
    Do
        If Trim(r) = "ZA" Then
            
            ' Range("A3:W3").Interior.Color = RGB(217, 217, 217)
            Range("A" & CStr(r.Row) & ":" & "W" & CStr(r.Row)).Interior.Color = RGB(217, 217, 217)
                
        End If
        Set r = r.Offset(1, 0)
    Loop While Trim(r.Value) <> ""
End Sub

Private Sub FMA_columns_adj()
    Columns("A:A").ColumnWidth = 3.71
    Columns("B:B").ColumnWidth = 8.71
    Columns("C:C").ColumnWidth = 10.29
    Columns("D:D").ColumnWidth = 32.14
    Columns("D:D").Font.size = 8
    Columns("E:F").ColumnWidth = 8.43
    Columns("G:G").ColumnWidth = 6.29
    Columns("H:H").ColumnWidth = 10.71
    Columns("I:I").ColumnWidth = 7.71
    Columns("K:P").ColumnWidth = 8.43
    Columns("I:I").Font.size = 9
    Columns("J:J").ColumnWidth = 9
    Columns("K:P").ColumnWidth = 8.43
    Columns("L:N").Font.size = 9
    Columns("O:P").Font.size = 8
    Columns("Q:R").ColumnWidth = 6.71
    Columns("S:S").ColumnWidth = 10.29
    Columns("T:U").ColumnWidth = 4.86
End Sub

Private Sub FMA_legenda()

      ' warunek if pod dodanie nowego wiersza
    If Cells(1, 1).Value = "PLT" Then
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
     
     ' kolor legendy
        Cells(1, 1).Interior.Color = RGB(255, 93, 93)
        Cells(1, 4).Interior.Color = RGB(164, 199, 57)
        Cells(1, 5).Interior.Color = RGB(217, 217, 217)
        Cells(1, 8).Interior.Color = RGB(164, 199, 57)
        Cells(1, 9).Interior.Color = RGB(164, 199, 57)
    End If

     'scalenie komorek legendy
    Range("A1:B1").Merge
    Range("F1:G1").Merge
    Range("I1:U1").Merge
    
     ' tresc legendy
    Cells(1, 3).Value = "NOK"
    Cells(1, 6).Value = "ZA"
    Cells(1, 9).Value = "FMA SCENARIO"
    
    'czcionka legendy
    Range("C1:U1").Font.Bold = True
    Range("C1:U1").Font.size = 14
    
    ' wysrodkowanie
    With Range("C1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Range("F1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Range("I1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
     'uklad legendy
    Dim rng As Range

    Set rng = Range("A1:C1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
        
    Set rng = Range("E1:G1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
    
   
    
End Sub

Private Sub component_leg()
  'scalenie komorek legendy
    Range("A1:B1").Merge
    Range("C1:D1").Merge
    Range("E1:V1").Merge
    
     ' tresc legendy
    Cells(1, 3).Value = "NOK"
    Cells(1, 5).Value = "COMPONENT SCENARIO"
    
    'czcionka legendy
    Range("C1:U1").Font.Bold = True
    Range("C1:U1").Font.size = 14
    Range("E1").Font.Color = RGB(0, 0, 0)
    
    ' wysrodkowanie
    With Range("C1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Range("E1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
     'uklad legendy
    Dim rng As Range

    Set rng = Range("A1:D1")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThick
    End With
        

    
        ' kolor legendy
    Cells(1, 1).Interior.Color = RGB(255, 0, 0)
    Cells(1, 5).Interior.Color = RGB(164, 199, 57)
End Sub

Private Sub component_columns()
    Columns("A:A").ColumnWidth = 3.71
    Columns("B:B").ColumnWidth = 8.71
    Columns("C:C").ColumnWidth = 10.29
    Columns("D:D").ColumnWidth = 32.14
    Columns("D:D").Font.size = 8
    Columns("E:H").ColumnWidth = 8.43
    Columns("I:I").ColumnWidth = 6.57
    Columns("J:K").ColumnWidth = 8.43
    Columns("G:N").Font.size = 8
    Columns("L:L").ColumnWidth = 6.29
    Columns("M:N").ColumnWidth = 7
    Columns("O:O").ColumnWidth = 7.71
    Columns("P:P").ColumnWidth = 10.71
    Columns("O:O").Font.size = 8
    Columns("Q:Q").ColumnWidth = 7.71
    Columns("Q:Q").Font.size = 8
    Columns("R:R").ColumnWidth = 9.29
    Columns("S:S").ColumnWidth = 7.29
    Columns("T:T").ColumnWidth = 9
    Columns("U:V").ColumnWidth = 4.86
End Sub

Private Sub ALL_columns_width()
    Columns("A:A").ColumnWidth = 3.71
    Columns("B:B").ColumnWidth = 8.71
    Columns("C:C").ColumnWidth = 10.29
    Columns("D:D").ColumnWidth = 17
    Columns("D:D").Font.size = 8
    Columns("E:E").ColumnWidth = 32.14
    Columns("E:G").Font.size = 8
    Columns("F:G").ColumnWidth = 13.57
    Columns("H:N").ColumnWidth = 7
    Columns("R:S").Font.size = 8
    Columns("O:U").ColumnWidth = 8.43
    Columns("W:W").ColumnWidth = 7
    Columns("W:W").Font.size = 8
    Columns("V:V").ColumnWidth = 7.71
    Columns("X:AI").ColumnWidth = 8.43
    Columns("AA:AE").Font.size = 8
    Columns("AJ:AQ").ColumnWidth = 9
    Columns("AR:AR").ColumnWidth = 20#
    Columns("AR:AR").Font.size = 8
    Columns("AS:AX").ColumnWidth = 7
    Columns("AY:AZ").ColumnWidth = 26.4
    Columns("AY:AZ").Font.size = 8
End Sub
