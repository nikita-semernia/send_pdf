Attribute VB_Name = "FormModerniserDevPPT"
Option Explicit

' This module should be loaded in PowerPoint only.

' Get the VBProject
Public Function VFM_ImportModule(ByVal stModulePath As String) As Object
  Set VFM_ImportModule = Application.ActivePresentation.VBProject.VBComponents.Import(stModulePath)
End Function

Public Function VFM_RemoveModule(ByVal stModuleName As String)
  With Application.ActivePresentation.VBProject
    On Error Resume Next
      .VBComponents.Remove .VBComponents(stModuleName)
    On Error GoTo 0
  End With
End Function

Public Function VFM_ExportModules(ByVal stModuleNames As String, ByVal stFolderPath As String)

  Const VBEXT_CT_STDMODULE = 1
  Const VBEXT_CT_CLASSMODULE = 2
  Const VBEXT_CT_MSFORM = 3
  
  Dim cmpComponent
  Dim stFileName As String

  stModuleNames = " " & stModuleNames & " "

  With Application.ActivePresentation.VBProject
    For Each cmpComponent In .VBComponents
      If InStr(stModuleNames, " " & cmpComponent.Name & " ") Then
        stFileName = vbNullString
        Select Case .VBComponents(cmpComponent.Name).Type
          Case VBEXT_CT_CLASSMODULE
            stFileName = cmpComponent.Name & ".cls"
          Case VBEXT_CT_MSFORM
            stFileName = cmpComponent.Name & ".frm"
          Case VBEXT_CT_STDMODULE
            stFileName = cmpComponent.Name & ".bas"
        End Select
        If stFileName <> vbNullString Then
          cmpComponent.Export VFMFileAddTrailingSlash(stFolderPath) & stFileName
        End If
      End If
    Next cmpComponent
    
  End With
End Function
