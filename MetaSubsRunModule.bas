Attribute VB_Name = "MetaSubsRunModule"
' ms7p3100
'ms7p3100_DEPT
'ms7p3100_OPER
'ms7p3100_PART_NAME
'ms7p3100_SCHED_PUBLISHED
'ms7p3100_SCHED_POINT
' =========================================================================================================

Private Sub refresh_ms7p3100(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
    Dim b_tmp As Boolean
    b_tmp = CBool(mh.m_ms7p3100.PN <> item.PN)
    b_tmp = CBool((mh.m_ms7p3100.PLT <> item.PLT) Or CBool(b_tmp))
    
    If b_tmp Then
        mh.m_ms7p3100.sPLT item.PLT
        mh.m_ms7p3100.sPN item.PN
        mh.m_ms7p3100.submit
    End If
        
End Sub


Public Sub ms7p3100_DEPT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p3100.NAME Then
        ' body of this subroutine
        refresh_ms7p3100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p3100.DEPT()
    Else
        mh.m_ms7p3100.open_this_screen
        ms7p3100_DEPT mh, col, item, osh
    End If
End Sub


Public Sub ms7p3100_OPER(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p3100.NAME Then
        ' body of this subroutine
        refresh_ms7p3100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p3100.OPER()
    Else
        mh.m_ms7p3100.open_this_screen
        ms7p3100_OPER mh, col, item, osh
    End If
End Sub


Public Sub ms7p3100_PART_NAME(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p3100.NAME Then
        ' body of this subroutine
        refresh_ms7p3100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p3100.PART_NAME()
    Else
        mh.m_ms7p3100.open_this_screen
        ms7p3100_PART_NAME mh, col, item, osh
    End If
End Sub


Public Sub ms7p3100_SCHED_PUBLISHED(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p3100.NAME Then
        ' body of this subroutine
        refresh_ms7p3100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p3100.SCHED_PUBLISHED()
    Else
        mh.m_ms7p3100.open_this_screen
        ms7p3100_SCHED_PUBLISHED mh, col, item, osh
    End If
End Sub


Public Sub ms7p3100_SCHED_POINT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p3100.NAME Then
        ' body of this subroutine
        refresh_ms7p3100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p3100.SCHED_POINT()
    Else
        mh.m_ms7p3100.open_this_screen
        ms7p3100_SCHED_POINT mh, col, item, osh
    End If
End Sub

' =========================================================================================================



' ms7p5100
'ms7p5100_DLY_COM
'ms7p5100_PLAN_COM
'ms7p5100_SUPP_NM
'ms7p5100_COUNTRY_CD
'ms7p5100_ZIP
'ms7p5100_DLY_XMIT
'ms7p5100_PLANNING_XMIT
'ms7p5100_FORECAST_FORMAT
'ms7p5100_ALIAS
'ms7p5100_ABRV_NM
' =========================================================================================================

Private Sub refresh_ms7p5100(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
        
    If (mh.m_ms7p5100.DUNS <> item.DUNS) Or (mh.m_ms7p5100.PLT <> item.PLT) Then
        mh.m_ms7p5100.sPLT item.PLT
        mh.m_ms7p5100.sDUNS item.DUNS
        mh.m_ms7p5100.submit
    End If
        
End Sub

Public Sub ms7p5100_DLY_COM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.DLY_COM()
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_DLY_COM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_PLAN_COM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.PLAN_COM()
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_PLAN_COM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_SUPP_NM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.SUPP_NM()
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_SUPP_NM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_COUNTRY_CD(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.COUNTRY_CD
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_COUNTRY_CD mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_ZIP(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.ZIP
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_ZIP mh, col, item, osh
    End If
End Sub


Public Sub ms7p5100_DLY_XMIT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.DLY_XMIT
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_DLY_XMIT mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_PLANNING_XMIT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.PLANNING_XMIT
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_PLANNING_XMIT mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_FORECAST_FORMAT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.FORECAST_FORMAT
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_FORECAST_FORMAT mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_ALIAS(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.ALIAS
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_ALIAS mh, col, item, osh
    End If
End Sub


Public Sub ms7p5100_ABRV_NM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.ABRV_NM
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_ABRV_NM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5100_SHIP_TIME_CODE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5100.NAME Then
        ' body of this subroutine
        refresh_ms7p5100 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5100.SHIP_TIME_CODE
    Else
        mh.m_ms7p5100.open_this_screen
        ms7p5100_SHIP_TIME_CODE mh, col, item, osh
    End If
End Sub

' =========================================================================================================







' ms7p5200
'ms7p5200_DESC
'ms7p5200_fieldNAME
'ms7p5200_STD_PACK
'ms7p5200_SCH_PACK
'ms7p5200_RECV_TYPE
'ms7p5200_FUT_RECV_TYPE
'ms7p5200_PAY_TYPE
'ms7p5200_BEGIN_DATE
'ms7p5200_PSUP_TYPE_P_S
' =========================================================================================================

Private Sub refresh_ms7p5200(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
        
    If (mh.m_ms7p5200.DUNS <> item.DUNS) Or _
        (mh.m_ms7p5200.PLT <> item.PLT) Or _
        (mh.m_ms7p5200.PN <> item.PN) Then
            mh.m_ms7p5200.sPLT item.PLT
            mh.m_ms7p5200.sDUNS item.DUNS
            mh.m_ms7p5200.sPN item.PN
            mh.m_ms7p5200.submit
    End If
        
End Sub

Public Sub ms7p5200_DESC(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.DESC()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_DESC mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_fieldNAME(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.fieldNAME()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_fieldNAME mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_STD_PACK(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.STD_PACK()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_STD_PACK mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_SCH_PACK(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.SCH_PACK()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_SCH_PACK mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_RECV_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.RECV_TYPE()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_RECV_TYPE mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_FUT_RECV_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.FUT_RECV_TYPE()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_FUT_RECV_TYPE mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_PAY_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.PAY_TYPE()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_PAY_TYPE mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_BEGIN_DATE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.BEGIN_DATE()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_BEGIN_DATE mh, col, item, osh
    End If
End Sub

Public Sub ms7p5200_PSUP_TYPE_P_S(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5200.NAME Then
        ' body of this subroutine
        refresh_ms7p5200 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5200.PSUP_TYPE_P_S()
    Else
        mh.m_ms7p5200.open_this_screen
        ms7p5200_PSUP_TYPE_P_S mh, col, item, osh
    End If
End Sub

' =========================================================================================================




' ms7p5900
'ms7p5900_SUPP_NM
'ms7p5900_NM
'ms7p5900_PRIM
'ms7p5900_SEC
'ms7p5900_MODE
'ms7p5900_TRANSIT
'ms7p5900_TC
' =========================================================================================================
Private Sub refresh_ms7p5900(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
        
    If (mh.m_ms7p5900.DUNS <> item.DUNS) Or (mh.m_ms7p5900.PLT <> item.PLT) Then
        mh.m_ms7p5900.sDUNS item.DUNS
        mh.m_ms7p5900.sPLT item.PLT
        mh.m_ms7p5900.submit
    End If
        
End Sub

Public Sub ms7p5900_SUPP_NM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.SUPP_NM()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_SUPP_NM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5900_NM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.NM()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_NM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5900_PRIM(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.PRIM()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_PRIM mh, col, item, osh
    End If
End Sub

Public Sub ms7p5900_SEC(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.SEC()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_SEC mh, col, item, osh
    End If
End Sub

Public Sub ms7p5900_MODE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.MODE()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_MODE mh, col, item, osh
    End If
End Sub

Public Sub ms7p5900_TRANSIT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.TRANSIT()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_TRANSIT mh, col, item, osh
    End If
End Sub

Public Sub ms7p5900_TC(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms7p5900.NAME Then
        ' body of this subroutine
        refresh_ms7p5900 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms7p5900.TC()
    Else
        mh.m_ms7p5900.open_this_screen
        ms7p5900_TC mh, col, item, osh
    End If
End Sub

' =========================================================================================================




' ms9pv400
'ms9pv400_STD_PACK
'ms9pv400_CNTR
'ms9pv400_GR_CONT_WG
'ms9pv400_UNIT_WT
' =========================================================================================================

Private Sub refresh_ms9pv400(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
        
    If (mh.m_ms9pv400.PLT <> item.PLT) Or (mh.m_ms9pv400.PN <> item.PN) Or (mh.m_ms9pv400.DUNS <> item.DUNS) Then
        mh.m_ms9pv400.sDUNS item.DUNS
        mh.m_ms9pv400.sKANBAN "    "
        mh.m_ms9pv400.sPLT item.PLT
        mh.m_ms9pv400.sPN item.PN
        mh.m_ms9pv400.submit
    End If
        
End Sub

Public Sub ms9pv400_STD_PACK(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pv400.NAME Then
        ' body of this subroutine
        refresh_ms9pv400 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pv400.STD_PACK()
    Else
        mh.m_ms9pv400.open_this_screen
        ms9pv400_STD_PACK mh, col, item, osh
    End If
End Sub


Public Sub ms9pv400_CNTR(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pv400.NAME Then
        ' body of this subroutine
        refresh_ms9pv400 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pv400.CNTR()
    Else
        mh.m_ms9pv400.open_this_screen
        ms9pv400_CNTR mh, col, item, osh
    End If
End Sub

Public Sub ms9pv400_GR_CONT_WG(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pv400.NAME Then
        ' body of this subroutine
        refresh_ms9pv400 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pv400.GR_CONT_WG()
    Else
        mh.m_ms9pv400.open_this_screen
        ms9pv400_GR_CONT_WG mh, col, item, osh
    End If
End Sub

Public Sub ms9pv400_UNIT_WT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pv400.NAME Then
        ' body of this subroutine
        refresh_ms9pv400 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pv400.UNIT_WT()
    Else
        mh.m_ms9pv400.open_this_screen
        ms9pv400_UNIT_WT mh, col, item, osh
    End If
End Sub


' =========================================================================================================



'mysprh40
'mysprh40_ROUTE1
'mysprh40_DOCK1
' =========================================================================================================
Private Sub refresh_mysprh40(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
        
    If (mh.m_mysprh40.PLT <> item.PLT) Or (mh.m_mysprh40.PN <> item.PN) Then
        mh.m_mysprh40.sPLT item.PLT
        mh.m_mysprh40.sPN item.PN
        mh.m_mysprh40.submit
    End If
        
End Sub

Public Sub mysprh40_ROUTE1(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_mysprh40.NAME Then
        ' body of this subroutine
        refresh_mysprh40 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_mysprh40.ROUTE1()
    Else
        mh.m_mysprh40.open_this_screen
        mysprh40_ROUTE1 mh, col, item, osh
    End If
End Sub

Public Sub mysprh40_DOCK1(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_mysprh40.NAME Then
        ' body of this subroutine
        refresh_mysprh40 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_mysprh40.DOCK1()
    Else
        mh.m_mysprh40.open_this_screen
        mysprh40_DOCK1 mh, col, item, osh
    End If
End Sub
' =========================================================================================================




' ms9pop00
'ms9pop00_BANK
'ms9pop00_A
'ms9pop00_F_U
'ms9pop00_DK
'ms9pop00_PCS_TO_GO
' =========================================================================================================
Private Sub refresh_ms9pop00(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
    If (mh.m_ms9pop00.PLT() <> item.PLT) Or (mh.m_ms9pop00.PN() <> item.PN) Then
        
            mh.m_ms9pop00.sPLT item.PLT
            mh.m_ms9pop00.sPN item.PN
            mh.m_ms9pop00.sDS "6"
            mh.m_ms9pop00.submit
    End If
        
End Sub
' =========================================================================================================

Public Sub ms9pop00_TT(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pop00.NAME Then
        ' body of this subroutine
        refresh_ms9pop00 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pop00.TT
    Else
        mh.m_ms9pop00.open_this_screen
        ms9pop00_TT mh, col, item, osh
    End If
End Sub

Public Sub ms9pop00_BANK(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pop00.NAME Then
        ' body of this subroutine
        refresh_ms9pop00 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pop00.bank
    Else
        mh.m_ms9pop00.open_this_screen
        ms9pop00_BANK mh, col, item, osh
    End If
End Sub

Public Sub ms9pop00_A(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pop00.NAME Then
        ' body of this subroutine
        refresh_ms9pop00 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pop00.a
    Else
        mh.m_ms9pop00.open_this_screen
        ms9pop00_A mh, col, item, osh
    End If
End Sub

Public Sub ms9pop00_F_U(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pop00.NAME Then
        ' body of this subroutine
        refresh_ms9pop00 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pop00.F_U
    Else
        mh.m_ms9pop00.open_this_screen
        ms9pop00_F_U mh, col, item, osh
    End If
End Sub

Public Sub ms9pop00_DK(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pop00.NAME Then
        ' body of this subroutine
        refresh_ms9pop00 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pop00.dk
    Else
        mh.m_ms9pop00.open_this_screen
        ms9pop00_DK mh, col, item, osh
    End If
End Sub

Public Sub ms9pop00_PCS_TO_GO(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_ms9pop00.NAME Then
        ' body of this subroutine
        refresh_ms9pop00 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_ms9pop00.pcs_to_go
    Else
        mh.m_ms9pop00.open_this_screen
        ms9pop00_PCS_TO_GO mh, col, item, osh
    End If
End Sub

' =========================================================================================================





' zk7ptogl
'zk7ptogl_CHECK_FRST_CURRENT_TYPE
'zk7ptogl_CHECK_FUTURE_CURRENT_TYPE
'zk7ptogl_CHECK_FRST_MUL
' =========================================================================================================
Private Sub refresh_zk7ptogl(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
    If mh.m_zk7ptogl.PLT <> item.PLT Or _
        mh.m_zk7ptogl.PN <> item.PN Or _
        mh.m_zk7ptogl.DUNS <> item.DUNS Then
        
            mh.m_zk7ptogl.sPLT item.PLT
            mh.m_zk7ptogl.sPN item.PN
            mh.m_zk7ptogl.sDUNS item.DUNS
            mh.m_zk7ptogl.submit
    End If
        
End Sub
' =========================================================================================================
Public Sub zk7ptogl_CHECK_FRST_CURRENT_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_zk7ptogl.NAME Then
        ' body of this subroutine
        refresh_zk7ptogl mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_zk7ptogl.CHECK_FRST_CURRENT_TYPE
    Else
        mh.m_zk7ptogl.open_this_screen
        zk7ptogl_CHECK_FRST_CURRENT_TYPE mh, col, item, osh
    End If
End Sub

Public Sub zk7ptogl_CHECK_FUTURE_CURRENT_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_zk7ptogl.NAME Then
        ' body of this subroutine
        refresh_zk7ptogl mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_zk7ptogl.CHECK_FUTURE_CURRENT_TYPE
    Else
        mh.m_zk7ptogl.open_this_screen
        zk7ptogl_CHECK_FUTURE_CURRENT_TYPE mh, col, item, osh
    End If
End Sub


Public Sub zk7ptogl_CHECK_FRST_MUL(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_zk7ptogl.NAME Then
        ' body of this subroutine
        refresh_zk7ptogl mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_zk7ptogl.CHECK_FRST_MUL
    Else
        mh.m_zk7ptogl.open_this_screen
        zk7ptogl_CHECK_FRST_MUL mh, col, item, osh
    End If
End Sub

' =========================================================================================================



' mysptog0
'mysptog0_CHECK_FRST_CURRENT_TYPE
'mysptog0_CHECK_FUTURE_CURRENT_TYPE
'mysptog0_CHECK_FRST_MUL
' =========================================================================================================
Private Sub refresh_mysptog0(ByRef mh As MgoHandler, ByRef item As InputitemHandler)
    
    If mh.m_mysptog0.PLT <> item.PLT Or _
        mh.m_mysptog0.PN <> item.PN Or _
        mh.m_mysptog0.DUNS <> item.DUNS Then
        
            mh.m_mysptog0.sPLT item.PLT
            mh.m_mysptog0.sPN item.PN
            mh.m_mysptog0.sDUNS item.DUNS
            mh.m_mysptog0.submit
    End If
        
End Sub
' =========================================================================================================
Public Sub mysptog0_CHECK_FRST_CURRENT_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_mysptog0.NAME Then
        ' body of this subroutine
        refresh_mysptog0 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_mysptog0.CHECK_FRST_CURRENT_TYPE
    Else
        mh.m_mysptog0.open_this_screen
        mysptog0_CHECK_FRST_CURRENT_TYPE mh, col, item, osh
    End If
End Sub

Public Sub mysptog0_CHECK_FUTURE_CURRENT_TYPE(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_mysptog0.NAME Then
        ' body of this subroutine
        refresh_mysptog0 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_mysptog0.CHECK_FUTURE_CURRENT_TYPE
    Else
        mh.m_mysptog0.open_this_screen
        mysptog0_CHECK_FUTURE_CURRENT_TYPE mh, col, item, osh
    End If
End Sub


Public Sub mysptog0_CHECK_FRST_MUL(mh As MgoHandler, col As Integer, item As InputitemHandler, osh As Worksheet)
    If UCase(mh.m_mgo.this_screen) = mh.m_mysptog0.NAME Then
        ' body of this subroutine
        refresh_mysptog0 mh, item
        osh.Cells(Int(item.iih_id), col) = mh.m_mysptog0.CHECK_FRST_MUL
    Else
        mh.m_mysptog0.open_this_screen
        mysptog0_CHECK_FRST_MUL mh, col, item, osh
    End If
End Sub

' =========================================================================================================

