Attribute VB_Name = "modStudentsStore"
Option Explicit

' =============================================================================
' modStudentsStore
' Purpose:
'   Load/save students.json (UTF-8) using the path stored in HKCU settings.
'
' Data model:
'   students.json is a JSON array of objects with fields:
'     - id (String) e.g. "u001"
'     - name (String)
'     - chat_id (Number / integer)
'     - active (Boolean, optional; default True)
'     - note (String, optional)
' =============================================================================

Private Const JSON_PRETTY_WHITESPACE As Long = 2

Private Function studentsJsonPath() As String
    ' Central place to resolve the current students.json path.
    studentsJsonPath = LoadStudentsJsonPath()
End Function

'Public Function ReadUtf8Text(ByVal path As String) As String
'    ' Reads UTF-8 text via ADODB.Stream and strips UTF-8 BOM if present.
'    Dim stm As Object
'    Set stm = CreateObject("ADODB.Stream")
'
'    stm.Type = 2          ' adTypeText
'    stm.Charset = "utf-8"
'    stm.Open
'    stm.LoadFromFile path
'
'    Dim s As String
'    s = stm.ReadText
'
'    stm.Close
'
'    If Len(s) > 0 Then
'        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2)
'    End If
'
'    ReadUtf8Text = s
'End Function

'Public Sub WriteUtf8Text(ByVal path As String, ByVal text As String)
'    ' Writes UTF-8 text via ADODB.Stream (overwrites file).
'    Dim stm As Object
'    Set stm = CreateObject("ADODB.Stream")
'
'    stm.Type = 2          ' adTypeText
'    stm.Charset = "utf-8"
'    stm.Open
'    stm.WriteText text
'    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
'    stm.Close
'End Sub

Public Function LoadStudents() As Collection
    ' Loads students.json.
    ' Returns an empty collection if the file is missing or invalid.
    On Error GoTo Fail

    Dim jsonText As String
    jsonText = modTextIO.ReadUtf8Text(studentsJsonPath)

    Dim col As Collection
    Set col = JsonConverter.ParseJson(jsonText) ' JSON array -> Collection
    Set LoadStudents = col
    Exit Function

Fail:
    Dim emptyCol As New Collection
    Set LoadStudents = emptyCol
End Function

Public Function TryLoadStudents(ByRef students As Collection, ByRef errText As String) As Boolean
    On Error GoTo Fail

    errText = vbNullString
    Set students = Nothing

    Dim path As String
    path = LoadStudentsJsonPath()

    If Len(Trim$(path)) = 0 Then
        errText = "Не задано шлях до students.json."
        GoTo SoftFail
    End If

    If Dir$(path) = vbNullString Then
        errText = "Файл students.json не знайдено:" & vbCrLf & path
        GoTo SoftFail
    End If

    Dim jsonText As String
    jsonText = modTextIO.ReadUtf8Text(path)

    Dim col As Collection
    Set col = JsonConverter.ParseJson(jsonText)

    Set students = col
    TryLoadStudents = True
    Exit Function

SoftFail:
    Set students = New Collection
    TryLoadStudents = False
    Exit Function

Fail:
    errText = "Не вдалося прочитати або розібрати students.json." & vbCrLf & _
              "Деталі: " & Err.Description
    Set students = New Collection
    TryLoadStudents = False
End Function


Public Sub SaveStudents(ByVal students As Collection)
    ' Saves students.json (pretty printed).
    Dim jsonOut As String
    jsonOut = JsonConverter.ConvertToJson(students, Whitespace:=JSON_PRETTY_WHITESPACE)

    modTextIO.WriteUtf8Text studentsJsonPath, jsonOut & vbCrLf
End Sub

Public Function NextStudentId(ByVal students As Collection) As String
    ' Finds the next available id in the form u001, u002, ...
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")

    Dim st As Dictionary
    For Each st In students
        If st.Exists("id") Then used(CStr(st("id"))) = True
    Next st

    Dim n As Long
    n = 1

    Do
        Dim sid As String
        sid = "u" & Format$(n, "000")

        If Not used.Exists(sid) Then
            NextStudentId = sid
            Exit Function
        End If

        n = n + 1
    Loop
End Function

Public Function NzStr(ByVal v As Variant) As String
    If IsMissing(v) Then NzStr = vbNullString: Exit Function
    If IsNull(v) Then NzStr = vbNullString: Exit Function
    NzStr = CStr(v)
End Function

Public Function ContainsText(ByVal hay As String, ByVal needle As String) As Boolean
    If Len(needle) = 0 Then
        ContainsText = True
    Else
        ContainsText = (InStr(1, hay, needle, vbTextCompare) > 0)
    End If
End Function

