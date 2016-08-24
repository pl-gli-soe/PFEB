VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PTOGL"
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


Private i_mgo As MGO

Public NAME As String


Private Sub Class_Initialize()
    ' Set i_mgo = New MGO
    NAME = "ZK7PTOGL"
End Sub

Private Sub Class_Terminate()
    ' Set i_mgo = Nothing
End Sub

Public Sub connect_with_mgo(arg As MGO)
    Set i_mgo = arg
End Sub

Public Sub submit()
    i_mgo.send_keys "<Enter>"
End Sub

Public Sub open_this_screen()
    i_mgo.send_keys "<Clear>"
    i_mgo.send_keys "zk7ptogl <Enter>"
End Sub

Public Function PLT()
    PLT = CStr(i_mgo.get_string(4, 7, 2))
End Function

Public Sub sPLT(arg As String)
    i_mgo.put_string CStr(arg), 4, 7
End Sub

Public Function DUNS()
    DUNS = CStr(i_mgo.get_string(4, 45, 9))
End Function

Public Sub sDUNS(arg As String)
    arg = CStr(add_zeros(arg, DUNS_LEN))
    i_mgo.put_string CStr(arg), 4, 45
End Sub

Public Function PN()
    PN = CStr(i_mgo.get_string(4, 19, 8))
End Function

Public Sub sPN(arg As String)
    arg = CStr(add_zeros(arg, PN_LEN))
    i_mgo.put_string CStr(arg), 4, 19
End Sub

Public Function CHECK_FRST_PN()
    CHECK_FRST_PN = CStr(i_mgo.get_string(11, 7, 8))
End Function

Public Function CHECK_FRST_MUL()
    CHECK_FRST_MUL = CStr(i_mgo.get_string(11, 7, 1))
End Function

Public Function CHECK_FRST_CURRENT_TYPE()
    CHECK_FRST_CURRENT_TYPE = CStr(i_mgo.get_string(11, 23, 1))
End Function

Public Function CHECK_FUTURE_CURRENT_TYPE()
    CHECK_FUTURE_CURRENT_TYPE = CStr(i_mgo.get_string(11, 33, 1))
End Function

