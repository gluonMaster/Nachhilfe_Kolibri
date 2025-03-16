Attribute VB_Name = "VBA_Export_All"
Sub ExportAllVBA()
    Dim vbComp As Object ' VBComponent
    Dim vbProj As Object ' VBProject
    Dim exportPath As String

    ' Set the folder for export (change the path if necessary)
    exportPath = ThisWorkbook.Path & "\ExportedVBA\"
    
    ' Create the folder if it does not exist
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    ' Get access to the VBA project
    Set vbProj = ThisWorkbook.VBProject

    ' Loop through all components (modules, forms, classes)
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                vbComp.Export exportPath & vbComp.name & ".bas"
            Case 2 ' Class Module
                vbComp.Export exportPath & vbComp.name & ".cls"
            Case 3 ' UserForm
                vbComp.Export exportPath & vbComp.name & ".frm"
        End Select
    Next vbComp

    MsgBox "All modules have been exported to " & exportPath, vbInformation
End Sub

