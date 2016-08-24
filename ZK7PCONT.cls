VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PCONT"
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
Private ile_wierszy As Integer

Public NAME As String


Private Sub Class_Initialize()
    ' Set i_mgo = New MGO
    
    NAME = "ZK7PCONT"
    ' !!! tu kontrolujemy ilosc wierszy na ekranie zk7pcont
    ile_wierszy = 4
End Sub

Private Sub Class_Terminate()
    ' Set i_mgo = Nothing
End Sub

Public Sub connect_with_mgo(arg As MGO)
    Set i_mgo = arg
End Sub

Public Sub submit()
    i_mgo.send_keys "<Enter>"
    ' wait jest juz wsadzony w send keys
    ' i_mgo.wait_for_mgo
End Sub

Public Function go_down() As Boolean
    
    ' dzieki temu dwie pieczenie na jednym ruszcie mozemy go down wrzucic jednoczesnie jako warunek sprawdzajacy
    ' czy mozemy leciec z iteracja
    If i_mgo.get_string(22, 2, 5) = "R6102" Then
        i_mgo.send_keys "<pf8>"
        go_down = True
    Else
        go_down = False
    End If
End Function

Public Sub open_this_screen()
    i_mgo.send_keys "<Clear>"
    i_mgo.send_keys "zk7pcont <Enter>"
    ' wait jest juz wsadzony w send keys
    ' i_mgo.wait_for_mgo
End Sub


' funkcja zwaracajaca ile moze byc wierszy danych w tym ekranie
' ==========================================================
Public Function MAX_ROWS() As Integer
    MAX_ROWS = ile_wierszy
End Function

' inputs
' ==========================================================
Public Function PLT()
    PLT = CStr(i_mgo.get_string(3, 18, 2))
End Function

Public Sub sPLT(arg As String)
    i_mgo.put_string CStr(arg), 3, 18
End Sub

Public Function DUNS()
    DUNS = CStr(i_mgo.get_string(4, 46, 9))
End Function

Public Sub sDUNS(arg As String)
    arg = CStr(add_zeros(arg, DUNS_LEN))
    i_mgo.put_string CStr(arg), 4, 46
End Sub

Public Function PN()
    PN = CStr(i_mgo.get_string(4, 18, 8))
End Function

Public Sub sPN(arg As String)
    arg = CStr(add_zeros(arg, PN_LEN))
    i_mgo.put_string CStr(arg), 4, 18
End Sub
' ==========================================================


' plt pobierany z nizszej czesci erkanu
' i as integer zaczynamy od zera
' i = 0 -> zatem pierwsze rownanie bedzie rowne 10 + 3 * 0
' i = 1 -> zatem pierwsze rownanie bedzie rowne 10 + 3 * 1
Public Function iPLT(i As Integer)
    iPLT = CStr(i_mgo.get_string(10 + 3 * Int(i), 2, 2))
End Function

Public Function iDUNS(i As Integer)
    iDUNS = CStr(i_mgo.get_string(10 + 3 * Int(i), 14, 9))
End Function


' czy jest literka P?
Public Function iA(i As Integer)
    iA = CStr(i_mgo.get_string(10 + 3 * Int(i), 12, 1))
End Function

Public Function iPO_NO(i As Integer)
    iPO_NO = CStr(i_mgo.get_string(10 + 3 * Int(i), 24, 8))
End Function


Public Function iSUPPLIER_NAME(i As Integer)
    iSUPPLIER_NAME = Trim(CStr(i_mgo.get_string(11 + 3 * Int(i), 13, 20)))
End Function

Public Function iEFF(i As Integer)
    iEFF = CStr(i_mgo.get_string(10 + 3 * Int(i), 48, 8))
End Function

Public Function iEXP(i As Integer)
    iEXP = CStr(i_mgo.get_string(11 + 3 * Int(i), 48, 8))
End Function


Public Function iDATE_ADDED(i As Integer)
    iDATE_ADDED = CStr(i_mgo.get_string(12 + 3 * Int(i), 2, 8))
End Function


Public Function iBYR(i As Integer)
    iBYR = Trim(CStr(i_mgo.get_string(10 + 3 * Int(i), 5, 5)))
End Function
