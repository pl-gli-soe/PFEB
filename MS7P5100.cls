VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS7P5100"
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
    NAME = "MS7P5100"
End Sub

Private Sub Class_Terminate()
    ' Set i_mgo = Nothing
End Sub

Public Sub connect_with_mgo(arg As MGO)
    Set i_mgo = arg
End Sub

Public Function PLT()
    PLT = CStr(i_mgo.get_string(1, 16, 2))
End Function

Public Sub sPLT(arg As String)
    i_mgo.put_string CStr(arg), 1, 16
End Sub

Public Function DUNS()
    DUNS = CStr(i_mgo.get_string(5, 14, 9))
End Function

Public Sub sDUNS(arg As String)

    arg = CStr(add_zeros(arg, DUNS_LEN))

    i_mgo.put_string CStr(arg), 5, 14
End Sub

Public Sub open_this_screen()
    i_mgo.send_keys "<Clear>"
    i_mgo.send_keys "ms7p5100 <Enter>"
End Sub


Public Sub submit()
    i_mgo.send_keys "<Enter>"
End Sub

Public Function SUPP_NM()
    SUPP_NM = Trim(CStr(i_mgo.get_string(6, 10, 34)))
End Function

Public Function DLY_COM()
    ' Debug.Print i_mgo.get_string(19, 11, 30)
    DLY_COM = Trim(CStr(i_mgo.get_string(19, 11, 30)))
End Function

Public Function PLAN_COM()
    ' Debug.Print i_mgo.get_string(20, 11, 30)
    PLAN_COM = Trim(CStr(i_mgo.get_string(20, 11, 30)))
End Function

Public Function COUNTRY_CD()
    COUNTRY_CD = Trim(CStr(i_mgo.get_string(13, 14, 2)))
End Function

Public Function ZIP()
    ZIP = Trim(CStr(i_mgo.get_string(12, 31, 12)))
End Function

Public Function DLY_XMIT()
    DLY_XMIT = Trim(CStr(i_mgo.get_string(4, 59, 1)))
End Function

Public Function PLANNING_XMIT()
    PLANNING_XMIT = Trim(CStr(i_mgo.get_string(4, 78, 1)))
End Function

Public Function FORECAST_FORMAT()
    FORECAST_FORMAT = Trim(CStr(i_mgo.get_string(18, 69, 1)))
End Function

Public Function ALIAS()
    ALIAS = Trim(CStr(i_mgo.get_string(5, 32, 15)))
End Function

Public Function ABRV_NM()
    ABRV_NM = Trim(CStr(i_mgo.get_string(7, 10, 15)))
End Function

Public Function SHIP_TIME_CODE()
    SHIP_TIME_CODE = Trim(CStr(i_mgo.get_string(13, 69, 10)))
End Function

