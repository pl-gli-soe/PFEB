Attribute VB_Name = "ExportThisProjectMod"
' working great!
Global Const REPO_PATH = "C:\WORKSPACE\macros\PFE\PFEB\repo\"

Private Sub export_this_project()
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_StdModule Then
            txt = VBComp.NAME & ".bas"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            txt = VBComp.NAME & ".cls"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            txt = VBComp.NAME & ".frm"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        End If
         
    Next VBComp
    
    MsgBox "ready!"

End Sub


Private Sub import_this_project()
    
    
    remove_current_implementation
    
    
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Set objFSO = New Scripting.FileSystemObject
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each objFile In objFSO.GetFolder(XWiz.REPO_PATH).Files
        ' body
        ' ==============================================================
        
        If (objFSO.GetExtensionName(objFile.NAME) = "cls") Or _
            (objFSO.GetExtensionName(objFile.NAME) = "frm") Or _
            (objFSO.GetExtensionName(objFile.NAME) = "bas") Then
            VBComps.Import objFile.Path
        End If
        
        ' ==============================================================
    Next objFile
    
    MsgBox "ready!"

End Sub


Private Sub remove_current_implementation()
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_Document Then
            txt = VBComp.NAME
            Debug.Print txt & " zostaje"
            
        ElseIf VBComp.Type = vbext_ct_ActiveXDesigner Then
            txt = VBComp.NAME
            Debug.Print txt & " zostaje"

        ElseIf CStr(VBComp.NAME) = "ExportThisProjectMod" Then
            txt = VBComp.NAME
            Debug.Print txt & " zostaje"
        Else
            
            VBComps.Remove VBComp
        End If
         
    Next VBComp
    
    ' MsgBox "ready!"

End Sub
