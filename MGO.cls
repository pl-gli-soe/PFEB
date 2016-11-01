VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PFEP licencja - grounds VBA source codes
' Copyright (c) 2013 mateusz.milewski@gm.com
'
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Private Sessions As Object
Private System As Object
Private sess0 As Object



Public Function get_sess0()
    Set get_sess0 = sess0
End Function



Private Sub Class_Initialize()
    ' mgo_init
End Sub


    
Private Sub Class_Terminate()
    ' nothing_on_mgo
End Sub

Public Function convertToDateFromMgoLongDate_implementation_is_empty(long_date As String) As Date

End Function

Public Function convertToDateFromMgoShortDate_implementation_is_empty(short_date As String) As Date
    
End Function

Public Function convertToDateFromMS9PO400Date(ms9po400_date As String) As Date

    ' format daty dla scr MS9PO400 -    ' 27MR12
    ' -----------------------------------------------------------YEAR----------------------------------------MONTH-------------------------------------------------------YEAR-------------------------------
    convertToDateFromMS9PO400Date = CDate(CStr(2000 + Int(Right(ms9po400_date, 2))) & "-" & CStr(frommgomonth(Right(Left(ms9po400_date, 4), 2))) & "-" & CStr(Int(Left(ms9po400_date, 2))))
End Function


' przyjmuje cala date z formatowaniem takim jakie wystepuje miedzy innymi na ekranie ms9pop00
' nie ma problemu zeby uzyc tej funkcji w dowolnym miejscu o ile wlasnie mamy takie formatowanie daty
' wewnatrz defincji znajduje sie wykorzystana funkcja FromMGOMonth
' funkcja zwraca date w formatowaniu zgodnym ze std excelowym
Public Function convertToDateFromMS9POP00Date(ms9pop00_date As String) As Date


    convertToDateFromMS9POP00Date = Date

    ' format daty dla scr MS9POP00 -    ' 27.MR.12
    ' -----------------------------------------------------------YEAR----------------------------------------MONTH-------------------------------------------------------YEAR-------------------------------
    convertToDateFromMS9POP00Date = CDate(CStr(2000 + Int(Right(ms9pop00_date, 2))) & "-" & CStr(frommgomonth(Right(Left(ms9pop00_date, 5), 2))) & "-" & CStr(Int(Left(ms9pop00_date, 2))))
End Function

Public Function convertToDateFromZK7PDRQMDate(zk7pdrqm_date As String) As Date

' instrukcja warunkowa wynika z powodu braku roku na ekranie zk7pdrqm i przez to program musi sam sie domyslec
' ze faktycznie nastapila zmiana roku- sam rozwiazalem to tak ze zazwyczaj dzieje sie tak nie dobrze, gdy chcemy
' stworzyc teardown nachodzi to na nowy rok, ale my jako ludzie zobaczymy tylko i wylacznie ze zaczynamy z powrotem
' od stycznia... postanowi³em zatem zostawic na stale ten sam rok i w przyapdku gdy przejdziemy do stycznia
' i sprawdzimy warunkiem to tak jakbysmy sie cofneli w czasie- cofniecie w czasie algorytm potraktuje jako coœ osobliwego
' zgodnie z trywialnym zalozeniem doda plus jeden do aktualnego roku
'
'
'
' na dzien 4 maja 2012


    'Debug.Print CStr(Format(Now, "yyyy"))
    'Debug.Print CStr(FromMGOMonth(Right(zk7pdrqm_date, 2)))
    'Debug.Print CStr(Left(zk7pdrqm_date, 2))
    '
    '
    'Debug.Print CStr(CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2))))
 
    ' tutaj mamy troche malo pieknie gdzie zasieg ff jest na tyle maly ze 100 dni jest out of scope dzieki czemu wie ze przeskakujemy na nastepny dzien
    ' jesli miedzy poszczegolnymi dniami iteracji jest az taka rozbieznosc moze sie zdarzyc ze bedzie 200 down days to wtedy fireflake
    ' sie wysypie faktycznie :P
    Dim tmp_date As Date
    tmp_date = CDate(Format(Now, "yyyy-mm-dd")) - 100

    If tmp_date <= CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(frommgomonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2))) Then
        'Debug.Print CStr(Format(Now, "yyyy"))
        'Debug.Print CStr(FromMGOMonth(Right(zk7pdrqm_date, 2)))
        'Debug.Print CStr(Left(zk7pdrqm_date, 2))
        
        '--------------------------------------------------YEAR-------------------------------------MONTH-------------------------------------DAY--------------
        convertToDateFromZK7PDRQMDate = CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(frommgomonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2)))
    Else
        '--------------------------------------------------YEAR-------------------------------------MONTH---------------------------------------------DAY--------------
        convertToDateFromZK7PDRQMDate = CDate(CStr(Int(Format(Now, "yyyy") + 1)) & "-" & CStr(frommgomonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2)))
    End If
End Function

Public Function convertToDateFromZK7PPUS0Date(zk7ppus0_date As String) As Date
    Dim d As String
    d = Right(CStr(zk7ppus0_date), 4) & "-" & CStr(frommgomonth(Right(Left(CStr(zk7ppus0_date), 5), 2))) & "-" & Left(CStr(zk7ppus0_date), 2)
    convertToDateFromZK7PPUS0Date = CDate(Format(d, "yyyy-mm-dd"))
End Function

Public Function convertToDateFromZK7PPUM0Date(zk7ppum0_date As String) As Date
    Dim d As String
    d = CStr(Left(CStr(Format(Now, "yyyy")), 2)) & Right(CStr(zk7ppum0_date), 2) & "-" & CStr(frommgomonth(Right(Left(CStr(zk7ppum0_date), 5), 2))) & "-" & Left(CStr(zk7ppum0_date), 2)
    convertToDateFromZK7PPUM0Date = CDate(Format(d, "yyyy-mm-dd"))
End Function

Public Function convertToMgoLongDate(n_date As Date) As String
    convertToMgoLongDate = Format(n_date, "dd") + "." + mgomonth(n_date) + "." + Format(n_date, "yyyy")
End Function

Public Function convertToMgoShortDate(n_date As Date) As String
    convertToMgoShortDate = Format(n_date, "dd") + "." + mgomonth(n_date) + "." + Format(n_date, "yy")
End Function

Public Function frommgomonth(mgomonth As String)
    frommgomonth = from_MGO_month(mgomonth)
End Function

Public Function from_MGO_month(mgo_month As String)

    Select Case mgo_month
        Case "JA"
            from_MGO_month = 1
        Case "FE"
            from_MGO_month = 2
        Case "MR"
            from_MGO_month = 3
        Case "AP"
            from_MGO_month = 4
        Case "MY"
            from_MGO_month = 5
        Case "JN"
            from_MGO_month = 6
        Case "JL"
            from_MGO_month = 7
        Case "AU"
            from_MGO_month = 8
        Case "SE"
            from_MGO_month = 9
        Case "OC"
            from_MGO_month = 10
        Case "NO"
            from_MGO_month = 11
        Case "DE"
            from_MGO_month = 12
    End Select
End Function

Public Function get_string(x As Integer, y As Integer, size As Integer) As String
    get_string = sess0.screen.getString(x, y, size)
    wait_for_mgo
End Function

Public Sub put_string(str As String, x As Integer, y As Integer)
    sess0.screen.putString str, x, y
    wait_for_mgo
End Sub

Public Sub mgo_init()
    Set System = CreateObject("EXTRA.System")   ' Gets the system object
    Set Sessions = System.Sessions
    Set sess0 = System.activeSession
    
    wait_for_mgo
End Sub

Public Sub nothing_on_mgo()
    Set sess0 = Nothing
    Set Sessions = Nothing
    Set System = Nothing
    
End Sub

Public Function mgomonth(n_date As Date) As String
    mgomonth = mgo_month(n_date)
End Function

Public Function mgo_month(n_date As Date) As String
    Select Case Format(n_date, "m") - 1
        Case 0
            mgo_month = "JA"
        Case 1
            mgo_month = "FE"
        Case 2
            mgo_month = "MR"
        Case 3
            mgo_month = "AP"
        Case 4
            mgo_month = "MY"
        Case 5
            mgo_month = "JN"
        Case 6
            mgo_month = "JL"
        Case 7
            mgo_month = "AU"
        Case 8
            mgo_month = "SE"
        Case 9
            mgo_month = "OC"
        Case 10
            mgo_month = "NO"
        Case 11
            mgo_month = "DE"
    End Select
End Function

Public Sub send_keys(str As String)
    sess0.screen.sendKeys (CStr(str))
    wait_for_mgo
End Sub

Public Sub wait_for_mgo()
    Do
        DoEvents
    Loop Until sess0.screen.OIA.XStatus = 0
End Sub

Public Function this_screen() As String
    this_screen = CStr(get_string(1, 2, 8))
End Function
