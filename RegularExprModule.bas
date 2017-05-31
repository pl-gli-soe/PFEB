Attribute VB_Name = "RegularExprModule"
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


'Precedence Table:
'
'Order  Name                Representation
'1      Parentheses         ( )
'2      Multipliers         ? + * {m,n} {m, n}?
'3      Sequence & Anchors  abc ^ $
'4      Alternation         |
'Predefined Character Abbreviations:
'
'abr    same as       meaning
'\d     [0-9]         Any single digit
'\D     [^0-9]        Any single character that's not a digit
'\w     [a-zA-Z0-9_]  Any word character
'\W     [^a-zA-Z0-9_] Any non-word character
'\s     [ \r\t\n\f]   Any space character
'\S     [^ \r\t\n\f]  Any non-space character
'\n     [\n]          New line
'example 1: Run As macro
'
'The following example macro looks at the value in cell A1 to see if the first 1 or 2 characters are digits.
'If so, they are removed and the rest of the string is displayed.
'If not, then a box appears telling you that no match is found. Cell A1 values of 12abc will return abc,
'value of 1abc will return abc, value of abc123 will return "Not Matched" because the digits were not at the start of the string.
'
'Private Sub simpleRegex()
'    Dim strPattern As String: strPattern = "^[0-9]{1,2}"
'    Dim strReplace As String: strReplace = ""
'    Dim regEx As New RegExp
'    Dim strInput As String
'    Dim Myrange As Range
'
'    Set Myrange = ActiveSheet.Range("A1")
'
'    If strPattern <> "" Then
'        strInput = Myrange.Value'
'
'        With regEx
'            .Global = True
'            .MultiLine = True
'            .IgnoreCase = False
'            .Pattern = strPattern
'        End With
'
'        If regEx.test(strInput) Then
'            MsgBox (regEx.Replace(strInput, strReplace))
'        Else
'            MsgBox ("Not matched")
'        End If
'    End If
'End Sub
'Example 2: Run as an in-cell function
'
'This example is the same as example 1 but is setup to run as an in-cell function.
'To use, change the code to this:
'
'Function simpleCellRegex(Myrange As Range) As String
'    Dim regEx As New RegExp
'    Dim strPattern As String
'    Dim strInput As String
'    Dim strReplace As String
'    Dim strOutput As String
'
'
'   strPattern = "^[0-9]{1,3}"
'
'    If strPattern <> "" Then
'        strInput = Myrange.Value
'        strReplace = ""'
'
'        With regEx
'            .Global = True
'            .MultiLine = True
'            .IgnoreCase = False
'            .Pattern = strPattern
'        End With
'
'        If regEx.test(strInput) Then
'            simpleCellRegex = regEx.Replace(strInput, strReplace)
'        Else
'            simpleCellRegex = "Not matched"
'        End If
'    End If
'End Function
'Place your strings ("12abc") in cell A1. Enter this formula =simpleCellRegex(A1) in cell B1 and the result will be "abc".
'
'
'
'Example 3: Loop Through Range
'
'This example is the same as example 1 but loops through a range of cells.'
'
'Private Sub simpleRegex()
'    Dim strPattern As String: strPattern = "^[0-9]{1,2}"
'    Dim strReplace As String: strReplace = ""
'    Dim regEx As New RegExp
'    Dim strInput As String
'    Dim Myrange As Range
'
'    Set Myrange = ActiveSheet.Range("A1:A5")
'
'    For Each cell In Myrange
'        If strPattern <> "" Then
'            strInput = cell.Value
'
'            With regEx
'                .Global = True
'                .MultiLine = True
'                .IgnoreCase = False
'                .Pattern = strPattern
'            End With
'
'            If regEx.test(strInput) Then
'                MsgBox (regEx.Replace(strInput, strReplace))
'            Else
'                MsgBox ("Not matched")
 '           End If
'        End If
'    Next
'End Sub
'Example 4: Splitting apart different patterns
'
'This example loops through a range (A1, A2 & A3) and looks for a string starting with three digits followed by a single alpha character and then 4 numeric digits. The output splits apart the pattern matches into adjacent cells by using the ().  $1 represents the first pattern matched within the first set of ().
'
'Private Sub splitUpRegexPattern()
'    Dim regEx As New RegExp
'    Dim strPattern As String
'    Dim strInput As String
'    Dim strReplace As String
'    Dim Myrange As Range
'
'    Set Myrange = ActiveSheet.Range("A1:A3")
'
'    For Each c In Myrange
'        strPattern = "(^[0-9]{3})([a-zA-Z])([0-9]{4})"
'
'        If strPattern <> "" Then
'            strInput = c.Value
'            strReplace = "$1"
'
'            With regEx
'                .Global = True
'                .MultiLine = True
'                .IgnoreCase = False
'                .Pattern = strPattern
'            End With
'
'            If regEx.test(strInput) Then
'                c.Offset(0, 1) = regEx.Replace(strInput, "$1")
'                c.Offset(0, 2) = regEx.Replace(strInput, "$2")
 ''               c.Offset(0, 3) = regEx.Replace(strInput, "$3")
 '           Else
 '               c.Offset(0, 1) = "(Not matched)"
  '          End If
  '      End If
  '  Next
'End Sub




'String   Regex Pattern                  Explanation
'a1aaa    [a-zA-Z][0-9][a-zA-Z]{3}       Single alpha, single digit, three alpha characters
'a1aaa    [a-zA-Z]?[0-9][a-zA-Z]{3}      May or may not have preceeding alpha character
'a1aaa    [a-zA-Z][0-9][a-zA-Z]{0,3}     Single alpha, single digit, 0 to 3 alpha characters
'a1aaa    [a-zA-Z][0-9][a-zA-Z]*         Single alpha, single digit, followed by any number of alpha characters
'
'</i8>    \<\/[a-zA-Z][0-9]\>            Exact non-word character except any single alpha followed by any single digit

Global Const pattern_item = "([a-zA-Z0-9_ <>!@#$%^&*()-+\/\\]{3})"





Public Function match_com_code(this_com_code) As E_JAKI_WZORZEC_ZOSTAL_SPELNIONY


    match_com_code = E_NO_MATCH

    Dim reg_ex As New RegExp
    
    '' dla 11 ''
    ' ([a-zA-Z0-9_ <>!@#$%^&*()-+\/\\]{3})\s([a-zA-Z0-9_ <>!@#$%^&*()-+\/\\]{3})\s([a-zA-Z0-9_ <>!@#$%^&*()-+\/\\]{3})
    ' Global Const pattern_item = "([a-zA-Z0-9_ <>!@#$%^&*()-+\/\\]{3})"
    
    wzorzec_glowny_11_znakowy = pattern_item & "\s" & pattern_item & "\s" & pattern_item
    wzorzec_glowny_7_znakowy = pattern_item & "\s" & pattern_item
    wzorzec_glowny_3_znakowy = pattern_item
    
    
    If Trim(this_com_code) <> "" Then
    
        ' dla wzorca glownego
        If Len(this_com_code) = 11 Then
        
            With reg_ex
                .Pattern = wzorzec_glowny_11_znakowy
                .MultiLine = False
                .Global = True
                .IgnoreCase = False
            End With
            
            If reg_ex.test(this_com_code) Then
                match_com_code = E_11_PATTERN
            End If
        
        ' wzorzec dla 2 czesciowego com code
        ElseIf Len(this_com_code) = 7 Then
        
        
            With reg_ex
                .Pattern = wzorzec_glowny_7_znakowy
                .MultiLine = False
                .Global = True
                .IgnoreCase = False
            End With
            
            If reg_ex.test(this_com_code) Then
                match_com_code = E_07_PATTERN
            End If
        
        
        ' wzorzec dla 3 znakowego com code
        ElseIf Len(this_com_code) = 3 Then
        
            
            With reg_ex
                .Pattern = wzorzec_glowny_3_znakowy
                .MultiLine = False
                .Global = True
                .IgnoreCase = False
            End With
            
            If reg_ex.test(this_com_code) Then
                match_com_code = E_03_PATTERN
            End If
            
        Else
        
            ' jeszcze cos innego
            match_com_code = E_NO_MATCH
        
        End If
    End If
End Function


Public Function match_with_tsk(this_com_code) As E_JAKI_WZORZEC_ZOSTAL_SPELNIONY
    Dim re As RegExp
    Set re = New RegExp
    
    
    ' Global Const pattern_item = "([a-zA-Z0-9_ <>!@#$%^&*()-+\/\\]{3})"
    
    ' match with tsk
    match_with_tsk = E_NO_MATCH
    
    ' czy w ogole jest co sprawdzac!
    If this_com_code Like "*TSK*" Then
    
    
        If PFEB.RegularExprModule.match_com_code(this_com_code) = E_11_PATTERN Then
        
            w1 = "TSK\s" & pattern_item & "\s" & pattern_item
            w2 = pattern_item & "\sTSK" & "\s" & pattern_item
            w3 = pattern_item & "\s" & pattern_item & "\sTSK"
            
            re.Pattern = w1
            If re.test(this_com_code) Then
                match_with_tsk = match_with_tsk + E_TSK_3_PRE
                
            End If
            
            re.Pattern = w2
            If re.test(this_com_code) Then
                match_with_tsk = match_with_tsk + E_TSK_3_MID
                
            End If
            
            re.Pattern = w3
            If re.test(this_com_code) Then
                match_with_tsk = match_with_tsk + E_TSK_3_END
                
            End If
        
        ElseIf PFEB.RegularExprModule.match_com_code(this_com_code) = E_07_PATTERN Then
        
            w1 = "TSK\s" & pattern_item
            w2 = pattern_item & "\sTSK"
            
            
            re.Pattern = w1
            If re.test(this_com_code) Then
                match_with_tsk = match_with_tsk + E_TSK_2_PRE
            End If
            
            re.Pattern = w2
            If re.test(this_com_code) Then
                match_with_tsk = match_with_tsk + E_TSK_2_END
            End If
            
        
        ElseIf PFEB.RegularExprModule.match_com_code(this_com_code) = E_03_PATTERN Then
            match_with_tsk = E_TSK_ALONE
        End If
    
    End If
    

End Function

Public Function match_with_cgh(this_com_code) As E_JAKI_WZORZEC_ZOSTAL_SPELNIONY


    Dim re As RegExp
    Set re = New RegExp
    
    ' match with CGH
    match_with_cgh = E_NO_MATCH
    
    ' czy w ogole jest co sprawdzac!
    If this_com_code Like "*CGH*" Then
    
    
        If PFEB.RegularExprModule.match_com_code(this_com_code) = E_11_PATTERN Then
        
            w1 = "CGH\s" & pattern_item & "\s" & pattern_item
            w2 = pattern_item & "\sCGH" & "\s" & pattern_item
            w3 = pattern_item & "\s" & pattern_item & "\sCGH"
            
            re.Pattern = w1
            If re.test(this_com_code) Then
                match_with_cgh = match_with_cgh + E_CGH_3_PRE
            End If
            
            re.Pattern = w2
            If re.test(this_com_code) Then
                match_with_cgh = match_with_cgh + E_CGH_3_MID
            End If
            
            re.Pattern = w3
            If re.test(this_com_code) Then
                match_with_cgh = match_with_cgh + E_CGH_3_END
            End If
        
        ElseIf PFEB.RegularExprModule.match_com_code(this_com_code) = E_07_PATTERN Then
        
            w1 = "CGH\s" & pattern_item
            w2 = pattern_item & "\sCGH"
            
            
            re.Pattern = w1
            If re.test(this_com_code) Then
                match_with_cgh = match_with_cgh + E_CGH_2_PRE
            End If
            
            re.Pattern = w2
            If re.test(this_com_code) Then
                match_with_cgh = match_with_cgh + E_CGH_2_END
            End If
            
        
        ElseIf PFEB.RegularExprModule.match_com_code(this_com_code) = E_03_PATTERN Then
            match_with_cgh = E_CGH_ALONE
        End If
    
    End If
End Function

