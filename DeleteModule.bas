Attribute VB_Name = "DeleteModule"
Public Sub delete_this_sheet(ictrl As IRibbonControl)
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If (ActiveSheet.NAME Like "*preinput*") Or _
        (ActiveSheet.NAME Like "*input*") Or _
        (ActiveSheet.NAME Like "*register*") Or _
        (ActiveSheet.NAME Like "*config*") Then
        MsgBox "you can't delete this sheet!"
    Else
        ActiveSheet.Delete
    End If
    
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


Public Sub delete_all_sheets(ictrl As IRibbonControl)

    Application.EnableEvents = False

    ret = MsgBox("Are you sure?", vbQuestion + vbYesNo)
    If ret = vbYes Then
        Application.DisplayAlerts = False
        
        x = 1
        Do
            If (Sheets(x).NAME Like "*preinput*") Or _
                (Sheets(x).NAME Like "*input*") Or _
                (Sheets(x).NAME Like "*register*") Or _
                (Sheets(x).NAME Like "*config*") Then
                x = x + 1
            Else
                Sheets(x).Delete
            End If
        Loop Until x > Sheets.COUNT
        Application.DisplayAlerts = True
    End If
    
    
    Application.EnableEvents = True
End Sub
