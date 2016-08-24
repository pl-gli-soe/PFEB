VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColourHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lbls As Range

Public Sub make_labels(ByRef osh As Worksheet)
    
    Set lbls = osh.Range("A1")
    Set lbls = osh.Range(lbls, lbls.End(xlToRight))
    
    lbls.WrapText = True
    lbls.EntireRow.AutoFit
    lbls.Font.Bold = True
    lbls.Font.Color = RGB(255, 255, 255)
    lbls.Interior.Color = RGB(164, 199, 57)
    
End Sub
