VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MgoHandler"
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


Public m_ms7p3100 As MS7P3100
Public m_ms7p5100 As MS7P5100
Public m_ms7p5200 As MS7P5200
Public m_ms7p5900 As MS7P5900
Public m_ms9pv400 As MS9PV400
Public m_mysprh40 As MYSPRH40
' component like zk7ptogl
Public m_mysptog0 As MYSPTOG0
Public m_ms9pop00 As MS9POP00
Public m_zk7ptogl As ZK7PTOGL
Public m_zk7pcont As ZK7PCONT

Public m_mgo As MGO


Private Sub Class_Initialize()
    Set m_mgo = New MGO
    m_mgo.mgo_init
    Set m_ms7p3100 = New MS7P3100
    Set m_ms7p5100 = New MS7P5100
    Set m_ms7p5200 = New MS7P5200
    Set m_ms7p5900 = New MS7P5900
    Set m_ms9pv400 = New MS9PV400
    Set m_mysprh40 = New MYSPRH40
    Set m_mysptog0 = New MYSPTOG0
    Set m_ms9pop00 = New MS9POP00
    Set m_zk7ptogl = New ZK7PTOGL
    Set m_zk7pcont = New ZK7PCONT
    
    m_ms7p3100.connect_with_mgo m_mgo
    m_ms7p5100.connect_with_mgo m_mgo
    m_ms7p5200.connect_with_mgo m_mgo
    m_ms7p5900.connect_with_mgo m_mgo
    m_ms9pv400.connect_with_mgo m_mgo
    m_mysprh40.connect_with_mgo m_mgo
    m_mysptog0.connect_with_mgo m_mgo
    m_ms9pop00.connect_with_mgo m_mgo
    m_zk7ptogl.connect_with_mgo m_mgo
    m_zk7pcont.connect_with_mgo m_mgo
    
    
    
End Sub

Private Sub Class_Terminate()
    Set m_ms7p3100 = Nothing
    Set m_ms7p5100 = Nothing
    Set m_ms7p5200 = Nothing
    Set m_ms7p5900 = Nothing
    Set m_ms9pv400 = Nothing
    Set m_mysprh40 = Nothing
    Set m_mysptog0 = Nothing
    Set m_ms9pop00 = Nothing
    Set m_zk7ptogl = Nothing
    Set m_zk7pcont = Nothing
    
    Set m_mgo = Nothing
End Sub
