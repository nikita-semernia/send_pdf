Attribute VB_Name = "modSettings"
Option Explicit

' HKCU\Software\VB and VBA Program Settings\<APP_NAME>\<SECTION>\<KEY>
Private Const APP_NAME  As String = "TutorSendPdf"
Private Const APP_TITLE As String = "Відправка PDF у Telegram"

Private Const SECTION_PATHS As String = "Paths"
Private Const KEY_STUDENTS_JSON As String = "StudentsJsonPath"
Private Const DEFAULT_STUDENTS_JSON As String = "C:\pdf\students.json"

Private Const SECTION_EXPORT As String = "Export"
Private Const KEY_QUALITY_PROFILE As String = "QualityProfile"
Private Const DEFAULT_QUALITY_PROFILE As String = "4K (3840x2160)"

Public Function AppName() As String
    AppName = APP_NAME   ' where APP_NAME = "TutorSendPdf"
End Function

Public Function AppTitle() As String
    AppTitle = APP_TITLE
End Function

Public Function LoadStudentsJsonPath() As String
    LoadStudentsJsonPath = GetSetting(APP_NAME, SECTION_PATHS, KEY_STUDENTS_JSON, DEFAULT_STUDENTS_JSON)
End Function

Public Sub SaveStudentsJsonPath(ByVal path As String)
    SaveSetting APP_NAME, SECTION_PATHS, KEY_STUDENTS_JSON, path
End Sub

Public Function LoadQualityProfile() As String
    LoadQualityProfile = GetSetting(APP_NAME, SECTION_EXPORT, KEY_QUALITY_PROFILE, DEFAULT_QUALITY_PROFILE)
End Function

Public Sub SaveQualityProfile(ByVal profile As String)
    SaveSetting APP_NAME, SECTION_EXPORT, KEY_QUALITY_PROFILE, profile
End Sub

Public Function DefaultStudentsJsonPath() As String
    DefaultStudentsJsonPath = DEFAULT_STUDENTS_JSON
End Function

Public Function DefaultQualityProfile() As String
    DefaultQualityProfile = DEFAULT_QUALITY_PROFILE
End Function

