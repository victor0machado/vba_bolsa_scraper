Attribute VB_Name = "VersionControl"
Function GetModulesPath()
  GetModulesPath = Application.ActiveWorkbook.Path & "\modules\"
End Function

Sub SaveCodeModules()
  Dim i, name
  
  With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.Count
      If .VBComponents(i).CodeModule.CountOfLines > 0 Then
        name = .VBComponents(i).CodeModule.name
        .VBComponents(i).Export GetModulesPath() & name & ".vba"
      End If
    Next i
  End With
End Sub

Sub ImportCodeModules()
  Dim i, module_name
  
  With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.Count
      module_name = .VBComponents(i).CodeModule.name
      
      If module_name <> "VersionControl" Then
        If Right(module_name, 6) = "Macros" Then
          .VBComponents.Remove .VBComponents(module_name)
          .VBComponents.Import GetModulesPath() & module_name & ".vba"
        End If
      End If
    Next i
  End With
End Sub
