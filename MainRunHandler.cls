VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainRunHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Private mh As MgoHandler
Private preish As Worksheet
Private ish As Worksheet
Private osh As Worksheet
Private cfg As Worksheet
Private r As Range, ir As Range, mr As Range

Private sh As StatusHandler

Private input_items As Collection
Private item As InputitemHandler


Public Sub add_new_output_sheet()
    'On Error Resume Next
    'ThisWorkbook.Sheets("output").Delete
    
    Set osh = ThisWorkbook.Sheets.Add(After:=Worksheets(Worksheets.COUNT))
    On Error Resume Next
    osh.NAME = "output" & ThisWorkbook.Sheets.COUNT
    
End Sub

Public Sub start_(mmh As MgoHandler)

    Set mh = mmh
    
    Set mr = osh.Range("a1")
    
    fill_labels
    
    Set input_items = Nothing
    Set input_items = New Collection
    
    Set ir = ish.Range("a2")
    for_iih_id = 2
    
    Do
    
        ' jesli jest filtr nalozony na liste to nawet nie wchodz na ten item
        If Not ir.EntireRow.Hidden Then
        
            Set item = Nothing
            Set item = New InputitemHandler
            item.duns = ir.Offset(0, 2)
            item.iih_id = for_iih_id
            item.pn = ir.Offset(0, 1)
            item.plt = ir
            input_items.Add item
            for_iih_id = for_iih_id + 1
        End If
    
        Set ir = ir.Offset(1, 0)
    Loop Until Trim(ir) = ""
    
    
    If input_items.COUNT > 0 Then
    
    
        Set sh = New StatusHandler
        sh.init_statusbar input_items.COUNT
        sh.show
        For Each item In input_items
        
            
            osh.Cells(item.iih_id, 1) = item.plt
            osh.Cells(item.iih_id, 2) = item.pn
            osh.Cells(item.iih_id, 3) = item.duns
        
            ' mamy przygotowane wszystkie dane nalezy teraz przeleciec przez odpowienie ekrany
            ' i pobrac dane
            Set r = cfg.Range("A2")
            Do
                If Trim(r.Offset(0, 2)) <> "" Then
                    Run CStr(r) & "_" & CStr(r.Offset(0, 1)), mh, Int(r.Offset(0, 2)), item, osh
                End If
                
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            sh.progress_increase
        Next item
    End If
    
    sh.hide
    
    
    Dim ch As ColourHandler
    Set ch = New ColourHandler
    
    ch.make_labels osh
    
    osh.Activate
    Columns("A:AT").Select
    Selection.ColumnWidth = SZEROKOSC_KOLUMNY
    
    
    freeze_top_row
    
    osh.Range("A1").Select
    
    
End Sub

Private Sub fill_labels()
    ' mr is a1
    
    ' init plt, pn, duns
    For x = 1 To 3
        mr.Offset(0, x - 1) = preish.Cells(1, Int(x))
    Next x
    
    Set mr = mr.Offset(0, 3)
    
    ' next from config
    Set r = cfg.Range("b2")
    Do
        ' -----------------------------------------------
        If Trim(r.Offset(0, 1)) <> "" Then
            If Trim(r.Offset(0, 3)) <> "" Then
                osh.Cells(1, Int(r.Offset(0, 1))) = CStr(Trim(r.Offset(0, 3)))
            Else
                osh.Cells(1, Int(r.Offset(0, 1))) = CStr(r.Offset(0, -1)) & Chr(10) & CStr(r)
            End If
        End If
        ' -----------------------------------------------
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
End Sub

Private Sub Class_Initialize()
    Set preish = ThisWorkbook.Sheets("preinput")
    Set ish = ThisWorkbook.Sheets("input")
    Set cfg = ThisWorkbook.Sheets("config")
    Set r = preish.Range("b2")
    Set ir = ish.Range("a2")
End Sub

