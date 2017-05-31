VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS7P3100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PFEP licencja - grounds VBA soource codes
' Copyright (c) 2013 mateusz.milewski@gm.com
'
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Public NAME As String

Private i_mgo As MGO


Private Sub Class_Initialize()
    ' Set i_mgo = New MGO
    NAME = "MS7P3100"
End Sub

Private Sub Class_Terminate()
    ' Set i_mgo = Nothing
End Sub

Public Sub connect_with_mgo(arg As MGO)
    Set i_mgo = arg
End Sub

Public Function plt()
    plt = CStr(i_mgo.get_string(1, 21, 2))
End Function

Public Sub sPLT(arg As String)
    i_mgo.put_string CStr(arg), 1, 21
End Sub

Public Function pn()
    pn = CStr(i_mgo.get_string(6, 13, 8))
End Function

Public Sub sPN(arg As String)

    arg = CStr(add_zeros(arg, PN_LEN))

    i_mgo.put_string CStr(arg), 6, 13
End Sub

Public Sub open_this_screen()
    i_mgo.send_keys "<Clear>"
    i_mgo.send_keys "ms7p3100 <Enter>"
End Sub


Public Sub submit()
    i_mgo.send_keys "<Enter>"
End Sub

Public Function DEPT()
    DEPT = Trim(CStr(i_mgo.get_string(6, 57, 5)))
End Function

Public Function OPER()
    OPER = Trim(CStr(i_mgo.get_string(6, 77, 4)))
End Function

Public Function PART_NAME()
    PART_NAME = Trim(CStr(i_mgo.get_string(7, 15, 15)))
End Function

Public Function SCHED_PUBLISHED()
    SCHED_PUBLISHED = Trim(CStr(i_mgo.get_string(14, 46, 1)))
End Function

Public Function SCHED_POINT()
    SCHED_POINT = Trim(CStr(i_mgo.get_string(19, 14, 14)))
End Function

