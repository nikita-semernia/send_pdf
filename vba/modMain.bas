Attribute VB_Name = "modMain"
Option Explicit

' =============================================================================
' Project: ExportAndSendPDF (PowerPoint VBA)
' Purpose:
'   1) Ask the user to select a student + export quality profile (frmStudents).
'   2) Export all slides to PNG at a chosen pixel resolution.
'   3) Call an external Python script that builds a PDF and sends it to Telegram.
'
' Notes:
'   - Slide.Export width/height are in pixels (ScaleWidth / ScaleHeight). [MS docs]
'   - Settings are stored under HKCU using GetSetting/SaveSetting.
' =============================================================================

' =============================================================================
' Configuration / constants
' =============================================================================

'' Registry keys (HKCU\Software\VB and VBA Program Settings\{APP_NAME}\{SECTION}\{KEY})
'Private Const APP_NAME As String = "TutorSendPdf"
'Private Const SECTION_EXPORT As String = "Export"
'Private Const KEY_QUALITY_PROFILE As String = "QualityProfile"

'' Default export profile (used when no setting exists yet)
'Private Const DEFAULT_PROFILE As String = "4K (3840x2160)"

' Log key emitted by Python (line format: KEY=value)
Private Const LOG_KEY_PDF_SIZE_BYTES As String = "PDF_SIZE_BYTES"

' Export file naming
Private Const SLIDE_PNG_PREFIX As String = "slide_"
Private Const SLIDE_PNG_EXT As String = ".png"
Private Const SLIDE_PNG_DIGITS As Long = 3

'' ---- Paths settings (HKCU) ----
'Private Const SECTION_PATHS As String = "Paths"
'Private Const KEY_STUDENTS_JSON As String = "StudentsJsonPath"
'Private Const DEFAULT_STUDENTS_JSON As String = "C:\Notes\plugins\send_pdf\students.json"

'Public Function LoadStudentsJsonPath() As String
'    LoadStudentsJsonPath = GetSetting(APP_NAME, SECTION_PATHS, KEY_STUDENTS_JSON, DEFAULT_STUDENTS_JSON)
'End Function

'Public Sub SaveStudentsJsonPath(ByVal path As String)
'    SaveSetting APP_NAME, SECTION_PATHS, KEY_STUDENTS_JSON, path
'End Sub

' =============================================================================
' Public API (used by forms / ribbon)
' =============================================================================

Public Sub Ribbon_SendPdf(ByVal control As IRibbonControl)
    ExportAndSendPDF
End Sub

Public Sub ManageStudents(ByVal control As IRibbonControl)
    frmStudentsManage.Show vbModal
End Sub

Public Sub ExportAndSendPDF()
    On Error GoTo FatalError

    ' ---- 1) Ask user for student + quality profile ----
    Dim studentId As String
    Dim studentLabel As String
    Dim profile As String
    Dim captionComment As String

    If Not PromptStudentAndProfile(studentId, studentLabel, profile, captionComment) Then Exit Sub


    ' Persist the chosen profile for next run.
    SaveQualityProfile profile

    ' ---- 2) Resolve profile to pixel resolution ----
    Dim exportW As Long, exportH As Long
    GetProfileWH profile, exportW, exportH

    ' ---- 3) Export slides to a temp folder as PNG ----
    Dim slidesDir As String
    slidesDir = CreateTempFolder("ppt_slides_")
    ExportSlidesToPng slidesDir, exportW, exportH
    
    ' ---- 4) Export JSON path ----
    Dim studentsJsonPath As String
    studentsJsonPath = modSettings.LoadStudentsJsonPath()
    
    ' ---- 5) Export caption ----
    Dim caption As String
    caption = "Конспект заняття за " & Format$(Date, "dd.mm.yyyy")
    
    If Len(Trim$(captionComment)) > 0 Then
        caption = caption & vbLf & Trim$(captionComment)
    End If

    ' ---- 6) Run Python and capture output to a log ----
    Dim pythonPath As String, scriptPath As String, logPath As String
    pythonPath = "C:\Users\nikit\AppData\Local\Programs\Python\Python312\python.exe"
    scriptPath = "C:\Notes\plugins\send_pdf\send_pdf.py"
    logPath = Environ$("TEMP") & "\ppt_send_log.txt"

    Dim exitCode As Long
    exitCode = RunPythonSend(pythonPath, scriptPath, slidesDir, studentId, studentLabel, profile, studentsJsonPath, caption, logPath)

    ' ---- 7) Parse optional PDF size from the log ----
    Dim pdfSizeMb As Double
    pdfSizeMb = ReadPdfSizeMbFromLog(logPath)

    ' ---- 8) Show result ----
    If exitCode = 0 Then
        ' Success toast is shown by Python; keep VBA silent.
    Else
        MsgBox "Failed (exitCode=" & exitCode & ")." & vbCrLf & "Check log: " & logPath, vbExclamation, modSettings.AppTitle()
    End If


    Exit Sub

FatalError:
    MsgBox "VBA error: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

' =============================================================================
' Settings (HKCU)
' =============================================================================

'Public Function LoadQualityProfile() As String
'    ' GetSetting returns Default if the key does not exist yet.
'    LoadQualityProfile = GetSetting(APP_NAME, SECTION_EXPORT, KEY_QUALITY_PROFILE, DEFAULT_PROFILE)
'End Function

'Public Sub SaveQualityProfile(ByVal profile As String)
'    ' SaveSetting stores values in HKCU for the current Windows user.
'    SaveSetting APP_NAME, SECTION_EXPORT, KEY_QUALITY_PROFILE, profile
'End Sub

Public Sub GetProfileWH(ByVal profile As String, ByRef W As Long, ByRef H As Long)
    ' 16:9 presets in pixels for Slide.Export ScaleWidth/ScaleHeight.
    Select Case profile
        Case "Fast (1280x720)":  W = 1280: H = 720
        Case "HD (1920x1080)":   W = 1920: H = 1080
        Case "2K (2560x1440)":   W = 2560: H = 1440
        Case "4K (3840x2160)":   W = 3840: H = 2160
        Case Else:               W = 3840: H = 2160
    End Select
End Sub

' =============================================================================
' UI helpers
' =============================================================================

Private Function PromptStudentAndProfile(ByRef studentId As String, _
                                        ByRef studentLabel As String, _
                                        ByRef profile As String, _
                                        ByRef captionComment As String) As Boolean
    ' Returns False if user cancelled / did not select a student.
    Dim frm As frmStudents
    Set frm = New frmStudents
    frm.Show vbModal

    studentId = frm.SelectedStudentId
    studentLabel = frm.SelectedStudentLabel
    profile = frm.SelectedQualityProfile
    captionComment = frm.SelectedCaptionComment

    Unload frm

    If Len(studentId) = 0 Then
        PromptStudentAndProfile = False
        Exit Function
    End If

    If Len(profile) = 0 Then profile = LoadQualityProfile()
    PromptStudentAndProfile = True

End Function

' =============================================================================
' Export logic
' =============================================================================

Private Sub ExportSlidesToPng(ByVal slidesDir As String, ByVal exportW As Long, ByVal exportH As Long)
    Dim sld As Slide
    Dim i As Long

    i = 1
    For Each sld In ActivePresentation.Slides
        ' Slide.Export scale width/height are pixels.
        sld.Export slidesDir & "\" & SLIDE_PNG_PREFIX & Format$(i, String$(SLIDE_PNG_DIGITS, "0")) & SLIDE_PNG_EXT, _
                   "PNG", exportW, exportH
        i = i + 1

        ' Keep UI responsive on large decks.
        If (i Mod 5) = 0 Then DoEvents
    Next sld
End Sub

Private Function CreateTempFolder(ByVal prefix As String) As String
    ' Creates a unique temp folder and returns its full path.
    Dim path As String
    path = Environ$("TEMP") & "\" & prefix & Format$(Now, "yyyymmdd_hhnnss")

    MkDir path
    CreateTempFolder = path
End Function

' =============================================================================
' Python execution
' =============================================================================

Private Function QuoteArg(ByVal s As String) As String
    QuoteArg = """" & Replace$(s, """", """""") & """"
End Function

Private Sub AppendText(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Append As #f
    Print #f, text
    Close #f
End Sub

Private Function RunPythonSend(ByVal pythonExe As String, _
                              ByVal scriptPath As String, _
                              ByVal slidesDir As String, _
                              ByVal studentId As String, _
                              ByVal studentLabel As String, _
                              ByVal profile As String, _
                              ByVal studentsJsonPath As String, _
                              ByVal caption As String, _
                              ByVal logPath As String) As Long
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim pythonw As String
    pythonw = Replace$(pythonExe, "python.exe", "pythonw.exe")

    Dim cmd As String
    cmd = QuoteArg(pythonw) & " " & _
          QuoteArg(scriptPath) & " " & _
          QuoteArg(slidesDir) & " " & _
          QuoteArg(studentId) & " " & _
          QuoteArg("--student-label") & " " & QuoteArg(studentLabel) & " " & _
          QuoteArg("--profile") & " " & QuoteArg(profile) & " " & _
          QuoteArg("--students-json") & " " & QuoteArg(studentsJsonPath) & " " & _
          QuoteArg("--caption") & " " & QuoteArg(caption) & " " & _
          QuoteArg("--log-path") & " " & QuoteArg(logPath)


    RunPythonSend = wsh.Run(cmd, 0, True)
End Function




' =============================================================================
' Log parsing / messaging
' =============================================================================

Private Function ReadPdfSizeMbFromLog(ByVal logPath As String) As Double
    ReadPdfSizeMbFromLog = 0

    If Dir$(logPath) = "" Then Exit Function

    Dim logText As String
    logText = modTextIO.ReadUtf8Text(logPath)

    Dim bytesStr As String
    bytesStr = FindLogValue(logText, LOG_KEY_PDF_SIZE_BYTES)

    If Len(bytesStr) > 0 And IsNumeric(bytesStr) Then
        ReadPdfSizeMbFromLog = CDbl(bytesStr) / 1024# / 1024#
    End If
End Function

Private Function BuildSuccessMessage(ByVal studentLabel As String, _
                                     ByVal studentId As String, _
                                     ByVal profile As String, _
                                     ByVal exportW As Long, _
                                     ByVal exportH As Long, _
                                     ByVal pdfSizeMb As Double, _
                                     ByVal logPath As String) As String
    BuildSuccessMessage = _
        "Done: sent successfully." & vbCrLf & _
        "Student: " & studentLabel & " (" & studentId & ")" & vbCrLf & _
        "Profile: " & profile & " (" & exportW & "x" & exportH & ")" & vbCrLf & _
        "PDF size: " & IIf(pdfSizeMb > 0, Format$(pdfSizeMb, "0.00") & " MB", "n/a") & vbCrLf & _
        "Log: " & logPath
End Function


Private Function FindLogValue(ByVal logText As String, ByVal key As String) As String
    ' Finds lines like:
    '   KEY=value
    ' and returns value for the first matching KEY.
    Dim lines() As String
    lines = Split(logText, vbCrLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim s As String
        s = lines(i)
        If Left$(s, Len(key) + 1) = key & "=" Then
            FindLogValue = Mid$(s, Len(key) + 2)
            Exit Function
        End If
    Next i

    FindLogValue = vbNullString
End Function

' =============================================================================
' Styling helper (shared by forms)
' =============================================================================

Public Sub ApplyModernStyle(ByVal frm As Object)
    ' Lightweight "modern" styling without external dependencies:
    ' - White background
    ' - Segoe UI (if available)
    ' - Applies to all controls that expose the Font property
    Dim c As control

    frm.BackColor = RGB(255, 255, 255)

    On Error Resume Next
    frm.Font.Name = "Segoe UI"
    frm.Font.Size = 10
    On Error GoTo 0

    For Each c In frm.Controls
        On Error Resume Next
        c.Font.Name = "Segoe UI"
        c.Font.Size = 10
        On Error GoTo 0
    Next c
End Sub


