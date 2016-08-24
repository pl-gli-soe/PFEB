Attribute VB_Name = "ScreenOpenersModule"


Public Sub open_ms7p3100(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_ms7p3100.open_this_screen
    
    mh.m_ms7p3100.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_ms7p3100.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_ms7p3100.submit
End Sub

Public Sub open_ms7p5100(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_ms7p5100.open_this_screen
    
    mh.m_ms7p5100.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_ms7p5100.sDUNS CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 3))

    mh.m_ms7p5100.submit
End Sub

Public Sub open_ms7p5200(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_ms7p5200.open_this_screen
    
    mh.m_ms7p5200.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_ms7p5200.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_ms7p5200.sDUNS CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 3))
    mh.m_ms7p5200.submit
End Sub

Public Sub open_ms7p5900(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_ms7p5900.open_this_screen
    
    mh.m_ms7p5900.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_ms7p5900.sDUNS CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 3))
    mh.m_ms7p5900.submit
End Sub

Public Sub open_ms9pop00(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_ms9pop00.open_this_screen
    
    mh.m_ms9pop00.sDS "6"
    mh.m_ms9pop00.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_ms9pop00.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_ms9pop00.submit
End Sub


Public Sub open_ms9pv400(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_ms9pv400.open_this_screen
    
    mh.m_ms9pv400.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_ms9pv400.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_ms9pv400.sDUNS CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 3))
    mh.m_ms9pv400.submit
End Sub


Public Sub open_mysprh40(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_mysprh40.open_this_screen
    
    mh.m_mysprh40.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_mysprh40.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_mysprh40.submit
End Sub


Public Sub open_zk7pcont(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_zk7pcont.open_this_screen
    
    mh.m_zk7pcont.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_zk7pcont.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    ' poniewaz chce zobaczyc inne dunsy
    ' mh.m_zk7pcont.sDUNS CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 3))
    mh.m_zk7pcont.submit
End Sub


Public Sub open_zk7ptogl(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_zk7ptogl.open_this_screen
    
    mh.m_zk7ptogl.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_zk7ptogl.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_zk7ptogl.submit
End Sub

Public Sub open_mysptog0(ictrl As IRibbonControl)
    Dim mh As MgoHandler
    Set mh = New MgoHandler
    
    mh.m_mysptog0.open_this_screen
    
    mh.m_mysptog0.sPLT CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 1))
    mh.m_mysptog0.sPN CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 2))
    mh.m_mysptog0.sDUNS CStr(ActiveCell.Parent.Cells(ActiveCell.Row, 3))
    mh.m_mysptog0.submit
End Sub
