Attribute VB_Name = "DeleteModule"
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

Public Sub delete_this_sheet(ictrl As IRibbonControl)
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If (ActiveSheet.NAME Like "*preinput*") Or _
        (ActiveSheet.NAME Like "*input*") Or _
        (ActiveSheet.NAME Like "*register*") Or _
        (ActiveSheet.NAME Like "*config*") Then
        MsgBox "you can't delete this sheet!"
    Else
        ActiveSheet.Delete
    End If
    
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


Public Sub delete_all_sheets(ictrl As IRibbonControl)

    Application.EnableEvents = False

    ret = MsgBox("Are you sure?", vbQuestion + vbYesNo)
    If ret = vbYes Then
        Application.DisplayAlerts = False
        
        x = 1
        Do
            If (Sheets(x).NAME Like "*preinput*") Or _
                (Sheets(x).NAME Like "*input*") Or _
                (Sheets(x).NAME Like "*register*") Or _
                (Sheets(x).NAME Like "*config*") Then
                x = x + 1
            Else
                Sheets(x).Delete
            End If
        Loop Until x > Sheets.COUNT
        Application.DisplayAlerts = True
    End If
    
    
    Application.EnableEvents = True
End Sub
