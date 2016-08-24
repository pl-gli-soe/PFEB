VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9POP00"
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
Private sess0 As Object


Private Sub Class_Initialize()
    ' Set i_mgo = New MGO
    NAME = "MS9POP00"
End Sub


Public Sub submit()
    i_mgo.send_keys "<Enter>"
End Sub

Public Sub open_this_screen()
    i_mgo.send_keys "<Clear>"
    i_mgo.send_keys "ms9pop00 <Enter>"
End Sub

Public Sub connect_with_mgo(arg As MGO)
    Set i_mgo = arg
    Set sess0 = i_mgo.get_sess0()
End Sub

Public Sub sPN(arg As String)
    arg = CStr(add_zeros(arg, PN_LEN))
    sess0.screen.putString CStr(arg), 3, 22
End Sub

Public Function PN() As String
    
    PN = Trim(sess0.screen.getString(4, 22, 8))
End Function

Public Sub sPLT(arg As String)
    
    sess0.screen.putString CStr(arg), 3, 13
End Sub

Public Function PLT() As String
    PLT = Trim(sess0.screen.getString(4, 13, 2))
End Function

Public Sub sDS(arg As String)
    sess0.screen.putString CStr(arg), 2, 18
End Sub

Public Function DS() As String
    DS = Trim(sess0.screen.getString(2, 18, 1))
End Function

Public Property Let DUNS(arg As String)
    arg = CStr(add_zeros(arg, DUNS_LEN))
    sess0.screen.putString CStr(arg), 3, 62
End Property

Public Property Get DUNS() As String
    DUNS = Trim(sess0.screen.getString(7, 7, 9))
End Property

Public Property Get SUPPLIER() As String
    SUPPLIER = Trim(sess0.screen.getString(7, 17, 20))
End Property

Public Property Let F_U(arg As String)
    sess0.screen.putString CStr(arg), 3, 35
End Property

Public Property Get F_U() As String
    F_U = Trim(sess0.screen.getString(4, 35, 2))
End Property

Public Property Let a(arg As String)
    sess0.screen.putString CStr(arg), 3, 53
End Property

Public Property Get a() As String
    a = Trim(sess0.screen.getString(4, 53, 2))
End Property


' pierwszy doh
' ===============================================================
Public Property Let firstDOH(arg As String)
    sess0.screen.putString CStr(arg), 3, 43
End Property

Public Property Get firstDOH() As String
    firstDOH = Trim(sess0.screen.getString(4, 43, 3))
End Property
' ===============================================================

' drugi doh
' ===============================================================
Public Property Let secDOH(arg As String)
    sess0.screen.putString CStr(arg), 3, 47
End Property

Public Property Get secDOH() As String
    secDOH = Trim(sess0.screen.getString(3, 47, 3))
End Property
' ===============================================================

' get DOH
' ===============================================================
Public Function DOH()
    DOH = Trim(CStr(i_mgo.get_string(4, 43, 3)))
End Function
' ===============================================================

Public Property Get MISC() As String
    MISC = Trim(sess0.screen.getString(5, 73, 7))
End Property

Public Property Get OS() As String
    OS = Trim(sess0.screen.getString(6, 67, 7))
End Property

Public Property Get bank() As String
    bank = Trim(sess0.screen.getString(6, 54, 7))
End Property


Public Property Get BBAL() As String
    If Trim(sess0.screen.getString(5, 15, 1)) = "-" Then
        BBAL = "-" & Trim(sess0.screen.getString(5, 7, 8))
    Else
        BBAL = Trim(sess0.screen.getString(5, 7, 8))
    End If
End Property


Public Property Get CBAL() As String
    If Trim(sess0.screen.getString(5, 30, 1)) = "-" Then
        CBAL = "-" & Trim(sess0.screen.getString(5, 22, 8))
    Else
        CBAL = Trim(sess0.screen.getString(5, 22, 8))
    End If
End Property


Public Property Get pcs_to_go() As String
    pcs_to_go = Trim(sess0.screen.getString(6, 23, 7))
End Property

Public Property Get dk() As String
    dk = Trim(sess0.screen.getString(8, 6, 8))
End Property

Public Property Get MODE() As String
    MODE = Trim(sess0.screen.getString(8, 21, 2))
End Property

Public Property Get TT() As String
    TT = Trim(sess0.screen.getString(8, 31, 8))
End Property

Public Property Get MNPC() As String
    MNPC = Trim(sess0.screen.getString(9, 7, 7))
End Property

Public Property Get NCX() As String
    NCX = Trim(sess0.screen.getString(10, 6, 1))
End Property

Public Property Get OBS() As String
    OBS = Trim(sess0.screen.getString(10, 14, 1))
End Property

Public Property Get STD_PACK() As String
    STD_PACK = Trim(sess0.screen.getString(10, 30, 9))
End Property

Public Property Get oneJOB() As String
    oneJOB = Trim(sess0.screen.getString(12, 7, 11))
End Property

Public Property Get IP() As String
    IP = Trim(sess0.screen.getString(12, 27, 5))
End Property

Public Property Get COUNT() As String
    COUNT = Trim(sess0.screen.getString(14, 8, 45))
End Property

Public Property Get PART_NAME() As String
    PART_NAME = Trim(sess0.screen.getString(6, 2, 10))
End Property

Public Property Get QHD() As String
    QHD = Trim(sess0.screen.getString(5, 48, 7))
End Property

Public Property Get o() As String
    o = Trim(sess0.screen.getString(15, 4, 25))
End Property



Public Function transQTY(i As Integer) As String
    transQTY = Trim(sess0.screen.getString(17 + (Int(i) * 2), 2, 9))
End Function

Public Function transSDATE(i As Integer) As String
    transSDATE = Trim(sess0.screen.getString(17 + (Int(i) * 2), 18, 8))
End Function

Public Function transPCS(i As Integer) As String
    transPCS = Trim(sess0.screen.getString(17 + (Int(i) * 2), 27, 3))
End Function

Public Function transWGT(i As Integer) As String
    transWGT = Trim(sess0.screen.getString(17 + (Int(i) * 2), 31, 6))
End Function

Public Function transMD(i As Integer) As String
    transMD = Trim(sess0.screen.getString(17 + (Int(i) * 2), 38, 2))
End Function

Public Function transCD(i As Integer) As String
    transCD = Trim(sess0.screen.getString(17 + (Int(i) * 2), 41, 1))
End Function

Public Function transEDA(i As Integer) As String
    transEDA = Trim(sess0.screen.getString(17 + (Int(i) * 2), 45, 8))
End Function

Public Function transETA(i As Integer) As String
    transETA = Trim(sess0.screen.getString(17 + (Int(i) * 2), 54, 2)) & ":" & Trim(sess0.screen.getString(17 + (Int(i) * 2), 56, 2))
End Function

Public Function transST(i As Integer) As String
    transST = Trim(sess0.screen.getString(17 + (Int(i) * 2), 60, 1))
End Function

Public Function transBL(i As Integer) As String
    transBL = Trim(sess0.screen.getString(17 + (Int(i) * 2), 63, 9))
End Function

Public Function transROUTE(i As Integer) As String
    transROUTE = Trim(sess0.screen.getString(17 + (Int(i) * 2), 73, 7))
End Function

Public Function transCMNT(i As Integer) As String
    transCMNT = Trim(sess0.screen.getString(18 + (Int(i) * 2), 39, 30))
End Function

Public Function transDUNS(i As Integer) As String
    transDUNS = Trim(sess0.screen.getString(18 + (Int(i) * 2), 71, 9))
End Function

Public Function transCONT(i As Integer) As String
    transCONT = Trim(sess0.screen.getString(18 + (Int(i) * 2), 2, 35))
End Function

