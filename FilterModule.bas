Attribute VB_Name = "FilterModule"
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


Public Sub quick_selection_filter(ictrl As IRibbonControl)
    quick_selection_filter_inner
End Sub


Private Sub quick_selection_filter_inner()
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    
    With fh
    
        .filtruj_po_selekcji
    
    End With

    Set fh = Nothing
End Sub


Public Sub quick_clear(ictrl As IRibbonControl)
    inner_quick_clear
End Sub

Public Sub inner_quick_clear()
    
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    fh.ustaw_swiezy_filtr
    
    Set fh = Nothing
End Sub
