Attribute VB_Name = "WgeneralModule"
Public Sub import_from_lgeneral(itrcl As IRibbonControl)
    MsgBox "Patience...to be implemented soon :)"
    ' inner_importjfgwerj, e,fe,wf,we, f,
End Sub

Public Sub import_from_wgeneral(itrcl As IRibbonControl)
    inner_import
End Sub



Private Sub inner_import()
    
    Dim c As Collection, i As WGeneralItem, r As Range, cc_sh As Worksheet
    
    
    Set cc_sh = ThisWorkbook.Sheets(PFEB.G_CC_SH)
    
    Dim wgen As Workbook
    Set wgen = get_wgen()
    
    ' MsgBox wgen.Name
    
    If Not wgen Is Nothing Then
        
        Set c = New Collection
        Set r = wgen.ActiveSheet.Cells(2, 1)
        Do
            
            If check_this_row(cc_sh, wgen, r) Then
                Set i = New WGeneralItem
                i.plt = CStr(wgen.ActiveSheet.Cells(r.Row, 2))
                i.pn = CStr(wgen.ActiveSheet.Cells(r.Row, 3))
                i.duns = CStr(wgen.ActiveSheet.Cells(r.Row, 13))
                i.country_code = CStr(wgen.ActiveSheet.Cells(r.Row, 15))
                i.is_osea_country_code = True
                

                i.alloc = CLng(CStr(wgen.ActiveSheet.Cells(r.Row, 11)))

                
                c.Add i
            End If
            
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
    End If
    
    wyrzuc_kolekcje_do_arkusza c
    
    wgen.Close False
    
End Sub

Public Function get_wgen() As Workbook
    
    Dim wgen As Workbook
    
    Set wgen = Nothing
    
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    
    int_choice = Application.FileDialog(msoFileDialogOpen).show
    
    If int_choice <> 0 Then
        sciezka_do_file_u = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        
        If Trim(sciezka_do_file_u) <> "" Then
        
            Set wgen = Workbooks.Open(sciezka_do_file_u)
            
            Do
                DoEvents
            Loop While wgen Is Nothing
        End If
    End If
    
    Set get_wgen = wgen
    
End Function

Private Function check_this_row(cc_sh As Worksheet, wgen As Workbook, r As Range) As Boolean
    check_this_row = False
    
    cc = Trim(CStr(wgen.ActiveSheet.Cells(r.Row, 15)))
    
    
    If find_this_in_country_code_sh(CStr(cc), cc_sh) Then
        check_this_row = True
    End If
End Function

Private Function find_this_in_country_code_sh(cc As String, cc_sh As Worksheet) As Boolean
    find_this_in_country_code_sh = False
    
    Dim r As Range
    Set r = cc_sh.Range("B2")
    Do
    
        If Trim(r) = Trim(cc) Then
            If Trim(CStr(r.Offset(0, 3))) = "1" Then
                find_this_in_country_code_sh = True
                
            Else
                find_this_in_country_code_sh = False
                
            End If
            
            Exit Function
        End If
    
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function

Public Sub wyrzuc_kolekcje_do_arkusza(c As Collection)
    
    Dim i As WGeneralItem, ish As Worksheet, ir As Range
    Set ish = ThisWorkbook.Sheets(PFEB.G_PRE_INPUT_SH)
    
    
    ish.Range("A2:C100000").Clear
    
    Set ir = ish.Cells(2, 1)
    For Each i In c
    
        
        
        ir = i.plt
        ir.Offset(0, 1) = i.pn
        'ir.Offset(0, 2) = i.duns
        
        
        Set ir = ir.Offset(1, 0)
        
    Next i
End Sub
