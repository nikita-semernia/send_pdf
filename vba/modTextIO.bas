Attribute VB_Name = "modTextIO"
Option Explicit

Public Function ReadUtf8Text(ByVal path As String) As String
    On Error GoTo Fail

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2          ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path

    Dim s As String
    s = stm.ReadText

    stm.Close

    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2)
    End If

    ReadUtf8Text = s
    Exit Function

Fail:
    Err.Raise Err.Number, "modTextIO.ReadUtf8Text", Err.Description
End Function

Public Sub WriteUtf8Text(ByVal path As String, ByVal text As String)
    On Error GoTo Fail

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2          ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
    stm.Close
    Exit Sub

Fail:
    Err.Raise Err.Number, "modTextIO.WriteUtf8Text", Err.Description
End Sub

