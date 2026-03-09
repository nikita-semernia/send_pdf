Attribute VB_Name = "modVbaExport"
Option Explicit

Private Const EXPORT_DIR As String = "C:\Notes\plugins\send_pdf\vba"
Private Const LOG_FILE As String = "C:\Notes\plugins\send_pdf\logs\vba_export.log"
Private Const APP_TITLE As String = "SendPdfProject"

Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2
Private Const VBEXT_CT_MSFORM As Long = 3

Public Sub ExportAllVba()
    On Error GoTo EH

    Dim vbProj As Object
    Dim vbComp As Object
    Dim exportedCount As Long

    EnsureFolderRecursive EXPORT_DIR
    EnsureFolderRecursive GetParentFolder(LOG_FILE)

    AppendLog "=== Export started ==="
    AppendLog "Target folder: " & EXPORT_DIR

    ClearOldExports EXPORT_DIR

    Set vbProj = Application.VBE.ActiveVBProject

    For Each vbComp In vbProj.VBComponents
        Select Case CLng(vbComp.Type)
            Case VBEXT_CT_STDMODULE, VBEXT_CT_CLASSMODULE, VBEXT_CT_MSFORM
                ExportComponent vbComp, EXPORT_DIR
                exportedCount = exportedCount + 1
                AppendLog "Exported: " & vbComp.Name
            Case Else
                AppendLog "Skipped: " & vbComp.Name & " (Type=" & CStr(vbComp.Type) & ")"
        End Select
    Next vbComp

    AppendLog "Export completed. Count=" & CStr(exportedCount)

    MsgBox "VBA export completed." & vbCrLf & _
           "Exported: " & exportedCount & vbCrLf & _
           "Folder: " & EXPORT_DIR, _
           vbInformation, APP_TITLE
    Exit Sub

EH:
    AppendLog "ERROR " & Err.Number & ": " & Err.Description

    MsgBox "VBA export failed." & vbCrLf & _
           "Reason: " & Err.Description & vbCrLf & vbCrLf & _
           "Check Trust Center setting:" & vbCrLf & _
           "Trust access to the VBA project object model", _
           vbExclamation, APP_TITLE
End Sub

Private Sub ExportComponent(ByVal vbComp As Object, ByVal targetDir As String)
    Dim ext As String
    Dim outPath As String

    ext = GetComponentExtension(CLng(vbComp.Type))
    If Len(ext) = 0 Then Exit Sub

    outPath = targetDir & "\" & vbComp.Name & ext

    DeleteIfExists outPath
    If LCase$(ext) = ".frm" Then
        DeleteIfExists targetDir & "\" & vbComp.Name & ".frx"
    End If

    vbComp.Export outPath
End Sub

Private Function GetComponentExtension(ByVal compType As Long) As String
    Select Case compType
        Case VBEXT_CT_STDMODULE
            GetComponentExtension = ".bas"
        Case VBEXT_CT_CLASSMODULE
            GetComponentExtension = ".cls"
        Case VBEXT_CT_MSFORM
            GetComponentExtension = ".frm"
        Case Else
            GetComponentExtension = vbNullString
    End Select
End Function

Private Sub ClearOldExports(ByVal targetDir As String)
    DeleteByMask targetDir, "*.bas"
    DeleteByMask targetDir, "*.cls"
    DeleteByMask targetDir, "*.frm"
    DeleteByMask targetDir, "*.frx"
End Sub

Private Sub DeleteByMask(ByVal folderPath As String, ByVal mask As String)
    Dim f As String
    f = Dir$(folderPath & "\" & mask)
    Do While Len(f) > 0
        Kill folderPath & "\" & f
        f = Dir$
    Loop
End Sub

Private Sub DeleteIfExists(ByVal filePath As String)
    If Len(Dir$(filePath)) > 0 Then Kill filePath
End Sub

Private Sub AppendLog(ByVal message As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open LOG_FILE For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #ff
    On Error GoTo 0
End Sub

Private Function GetParentFolder(ByVal filePath As String) As String
    GetParentFolder = Left$(filePath, InStrRev(filePath, "\") - 1)
End Function

Private Sub EnsureFolderRecursive(ByVal folderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long

    If Len(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> "" Then Exit Sub

    parts = Split(folderPath, "\")
    currentPath = parts(0)

    If Right$(currentPath, 1) <> "\" Then currentPath = currentPath & "\"

    For i = 1 To UBound(parts)
        currentPath = currentPath & parts(i)
        If Dir$(currentPath, vbDirectory) = "" Then MkDir currentPath
        If i < UBound(parts) Then currentPath = currentPath & "\"
    Next i
End Sub


