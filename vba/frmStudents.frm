VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStudents 
   Caption         =   "Відправити PDF-копію у Telegram"
   ClientHeight    =   3936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6852
   OleObjectBlob   =   "frmStudents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' frmStudents
' Purpose:
'   Modal picker for:
'     - Student (display name + hidden id)
'     - Export quality profile (resolution preset)
'
' Outputs (read by caller after frm.Show vbModal):
'   - SelectedStudentId
'   - SelectedStudentLabel
'   - SelectedQualityProfile
'
' Notes:
'   - Student list is loaded from students.json (UTF-8) produced/maintained externally.
'   - The ComboBox stores 2 columns: [0]=display label, [1]=id (hidden column).
' =============================================================================

' Public outputs (read by modMain)
Public SelectedStudentId As String
Public SelectedStudentLabel As String
Public SelectedQualityProfile As String
Public SelectedCaptionComment As String

' vba-form-moderniser manager (turns CommandButtons into label-based "modern" buttons)
Private m_LabelControlsMgr As CLabelControlsManager

' -----------------------------------------------------------------------------
' Configuration
' -----------------------------------------------------------------------------

Private Const PLACEHOLDER_TEXT As String = "— Обрати отримувача —"
Private Const MSG_TITLE_REQUIRED As String = "Потрібно обрати"
Private Const MSG_PICK_RECIPIENT As String = "Треба обрати отримувача."

Private Const COL_LABEL As Long = 0
Private Const COL_ID As Long = 1

Public Sub btnStudentsSettings_Click()
    frmStudentsManage.Show vbModal
    ReloadStudentsList
End Sub

Private Sub ReloadStudentsList()
    Dim students As Collection
    Dim errText As String

    With Me.cmbStudents
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "300 pt;0 pt"
        .AddItem PLACEHOLDER_TEXT
        .List(.ListCount - 1, COL_ID) = vbNullString
        .ListIndex = 0
    End With

    If Not modStudentsStore.TryLoadStudents(students, errText) Then
        ' Keep picker open; just show error and leave placeholder.
        modUiDialogs.ShowError errText
        Exit Sub
    End If

    Dim st As Dictionary
    For Each st In students
        If StudentIsActive(st) Then AddStudentRow st
    Next st

    Me.cmbStudents.ListIndex = 0
End Sub


' -----------------------------------------------------------------------------
' Form lifetime
' -----------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    ' 1) Basic styling (font/background)
    ApplyModernStyle Me

    ' 2) Hook moderniser (buttons + general control styling)
    InitialiseModerniser

    ' 3) Keyboard defaults (may not override ComboBox behavior; we handle KeyDown too)
    Me.btnOK.Default = True
    Me.btnCancel.Cancel = True

    ' 4) Populate UI
    InitQualityProfiles
    InitStudents

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Removes hover highlight when mouse leaves a modernised button.
    If Not m_LabelControlsMgr Is Nothing Then
        m_LabelControlsMgr.LabelControls.UpdateControlButtonState
    End If
End Sub

' -----------------------------------------------------------------------------
' Moderniser bootstrap
' -----------------------------------------------------------------------------

Private Sub InitialiseModerniser()
    ' General restyling of MSForms controls
    ModerniseControls Me.Controls

    ' Configure moderniser buttons:
    FormModerniserModule.ActiveButton = vbNullString
    FormModerniserModule.DefaultButton = "btnOK"

    Dim tabOrder() As String
    tabOrder = Split("btnOK btnCancel")

    Set m_LabelControlsMgr = VFMFactory.CreateCLabelControlsManager(Me, Me.Controls, tabOrder)
End Sub


' -----------------------------------------------------------------------------
' UI init: quality profiles
' -----------------------------------------------------------------------------

Private Sub InitQualityProfiles()
    With Me.cmbQuality
        .Clear
        .AddItem "Fast (1280x720)"
        .AddItem "HD (1920x1080)"
        .AddItem "2K (2560x1440)"
        .AddItem "4K (3840x2160)"
    End With

    ' Restore previously used profile (fallback to 4K).
    Dim saved As String
    saved = LoadQualityProfile()

    Dim i As Long
    For i = 0 To Me.cmbQuality.ListCount - 1
        If Me.cmbQuality.List(i) = saved Then
            Me.cmbQuality.ListIndex = i
            Exit For
        End If
    Next i
    If Me.cmbQuality.ListIndex < 0 Then Me.cmbQuality.ListIndex = 3

    ' Show initial estimate
    UpdateApproxSize
End Sub

Private Sub cmbQuality_Change()
    UpdateApproxSize
End Sub

Private Sub UpdateApproxSize()
    ' Shows approximate multiplier vs 1080p.
    Dim W As Long, H As Long
    GetProfileWH CStr(Me.cmbQuality.value), W, H

    Dim ratio As Double
    ratio = (CDbl(W) * CDbl(H)) / (1920# * 1080#)

    Me.lblApproxSize.caption = "Approx size: " & Format$(ratio, "0.00x") & " vs HD"
End Sub

' -----------------------------------------------------------------------------
' UI init: students list
' -----------------------------------------------------------------------------

Private Sub InitStudents()
    ' Always initialise ComboBox (no crashes, even if data load fails)
    With Me.cmbStudents
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "300 pt;0 pt" ' hide id column

        .AddItem PLACEHOLDER_TEXT
        .List(.ListCount - 1, COL_ID) = vbNullString
        .ListIndex = 0
    End With

    Dim students As Collection
    Dim errText As String

    If Not modStudentsStore.TryLoadStudents(students, errText) Then
        Dim msg As String
        msg = errText & vbCrLf & vbCrLf & _
              "Відкрити «Налаштування списку учнів», щоб виправити шлях/дані?"

        If modUiDialogs.AskOpenStudentsManage(errText) Then
            ' Let user fix path/data, then try loading again (no nested Show vbModal!)
            frmStudentsManage.SuppressLoadErrorPopup = True
            frmStudentsManage.Show vbModal
            frmStudentsManage.SuppressLoadErrorPopup = False

        
            If Not modStudentsStore.TryLoadStudents(students, errText) Then
                ' Still broken after Manage
                modUiDialogs.ShowError errText
                Exit Sub
            End If
        Else
            Exit Sub
        End If

    End If

    Dim st As Dictionary
    For Each st In students
        If StudentIsActive(st) Then AddStudentRow st
    Next st

    Me.cmbStudents.ListIndex = 0
End Sub


Private Function StudentIsActive(ByVal st As Dictionary) As Boolean
    ' Default to active when field is missing.
    If st.Exists("active") Then
        StudentIsActive = CBool(st("active"))
    Else
        StudentIsActive = True
    End If
End Function

Private Sub AddStudentRow(ByVal st As Dictionary)
    Dim label As String
    label = CStr(st("name"))

    If st.Exists("note") Then
        If Len(CStr(st("note"))) > 0 Then
            label = label & " — " & CStr(st("note"))
        End If
    End If

    Me.cmbStudents.AddItem label
    Me.cmbStudents.List(Me.cmbStudents.ListCount - 1, COL_ID) = CStr(st("id"))
End Sub

' -----------------------------------------------------------------------------
' Buttons (public: required by moderniser CallByName)
' -----------------------------------------------------------------------------

Public Sub btnOK_Click()
    ' Validate selection (index 0 is the placeholder).
    If Me.cmbStudents.ListIndex <= 0 Then
        modUiDialogs.ShowError MSG_PICK_RECIPIENT
        Me.cmbStudents.SetFocus
        Exit Sub
    End If

    ' Persist & return selected profile
    SaveQualityProfile CStr(Me.cmbQuality.value)
    Me.SelectedQualityProfile = CStr(Me.cmbQuality.value)

    ' Return selected student id + label + comment
    Me.SelectedStudentId = CStr(Me.cmbStudents.List(Me.cmbStudents.ListIndex, COL_ID))
    Me.SelectedStudentLabel = CStr(Me.cmbStudents.List(Me.cmbStudents.ListIndex, COL_LABEL))
    Me.SelectedCaptionComment = Trim$(Me.txtCaptionComment.text)

    Me.Hide
End Sub

Public Sub btnCancel_Click()
    SelectedStudentId = vbNullString
    SelectedStudentLabel = vbNullString
    SelectedQualityProfile = vbNullString
    SelectedCaptionComment = vbNullString
    Me.Hide
End Sub

' -----------------------------------------------------------------------------
' Keyboard handling
' -----------------------------------------------------------------------------

Private Sub cmbStudents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleCommonKeys KeyCode

    ' Prevent focus from jumping to the next control when the user presses
    ' Up/Down at list boundaries (user wants to stay in this ComboBox).
    If KeyCode = vbKeyDown Then
        If Me.cmbStudents.ListIndex = Me.cmbStudents.ListCount - 1 Then KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        If Me.cmbStudents.ListIndex <= 0 Then KeyCode = 0
    End If
End Sub

Private Sub cmbQuality_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleCommonKeys KeyCode

    If KeyCode = vbKeyDown Then
        If Me.cmbQuality.ListIndex = Me.cmbQuality.ListCount - 1 Then KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        If Me.cmbQuality.ListIndex <= 0 Then KeyCode = 0
    End If
End Sub

Private Sub HandleCommonKeys(ByRef KeyCode As MSForms.ReturnInteger)
    ' Force consistent behavior for Enter/Escape regardless of focus.
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        btnOK_Click
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        btnCancel_Click
    End If
End Sub


