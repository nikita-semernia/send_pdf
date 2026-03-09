Attribute VB_Name = "modUiDialogs"
Option Explicit

' Title comes from modSettings (single source of truth)
Public Function AppTitle() As String
    AppTitle = modSettings.AppTitle()
End Function

' Shows an error with optional details (keeps UI text consistent)
Public Sub ShowError(ByVal message As String, Optional ByVal details As String = vbNullString)
    Dim s As String
    s = message
    If Len(details) > 0 Then s = s & vbCrLf & vbCrLf & details
    MsgBox s, vbExclamation, AppTitle()
End Sub

' Asks user whether to open student manager to fix path/JSON.
' Returns True if user chose Yes.
Public Function AskOpenStudentsManage(ByVal errText As String) As Boolean
    Dim s As String
    s = errText & vbCrLf & vbCrLf & _
        "Відкрити «Налаштування списку учнів», щоб виправити шлях/дані?"
    AskOpenStudentsManage = (MsgBox(s, vbYesNo + vbExclamation, AppTitle()) = vbYes)
End Function

Public Sub ShowInfo(ByVal message As String)
    MsgBox message, vbInformation, AppTitle()
End Sub


