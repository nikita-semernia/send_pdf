VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStudentsManage 
   Caption         =   "Налаштування списку учнів"
   ClientHeight    =   7008
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6804
   OleObjectBlob   =   "frmStudentsManage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStudentsManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' frmStudentsManage
' Purpose:
'   CRUD editor for students.json:
'     - Create/update student entries
'     - Deactivate / delete
'     - Search/filter (including optional inactive entries)
'     - Change students.json location (Browse...)
'
' UI notes:
'   - TextBox.Tag is used as placeholder text.
'   - Placeholder state is detected via ForeColor.
' =============================================================================

' In-memory model
Private m_students As Collection

' 1-based index in m_students (Collection); 0 means "no selection / new record"
Private m_selectedIndex As Long

' Placeholder color (light gray)
Private Const PLACEHOLDER_COLOR As Long = 9868950 ' RGB(150,150,150)

' ListBox columns
Private Const COL_LABEL As Long = 0
Private Const COL_ID As Long = 1

' Moderniser manager (turns CommandButtons into modern label-based buttons)
Private m_LabelControlsMgr As CLabelControlsManager

Public SuppressLoadErrorPopup As Boolean


' -----------------------------------------------------------------------------
' Form lifetime
' -----------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    m_selectedIndex = 0

    ' Match the look and feel of frmStudents
    ApplyModernStyle Me
    InitialiseModerniser

    ' Show current students.json path (read-only)
    Me.txtStudentsJsonPath.Locked = True
    Me.txtStudentsJsonPath.TabStop = False
    Me.txtStudentsJsonPath.text = LoadStudentsJsonPath()

    ' Placeholders (placeholder text is stored in each TextBox.Tag)
    ApplyPlaceholder txtSearch
    ApplyPlaceholder txtName
    ApplyPlaceholder txtChatId
    ApplyPlaceholder txtNote

    ReloadStudentsFromDisk
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
    ModerniseControls Me.Controls

    FormModerniserModule.ActiveButton = vbNullString
    FormModerniserModule.DefaultButton = "btnSave"

    Dim tabOrder() As String
    tabOrder = Split("btnNew btnSave btnDeactivate btnDelete btnBrowseStudentsJson btnClose")

    Set m_LabelControlsMgr = VFMFactory.CreateCLabelControlsManager(Me, Me.Controls, tabOrder)
End Sub

Private Sub SetCollectionSingle(ByVal col As Collection, ByVal obj As Object)
    Dim i As Long
    For i = col.Count To 1 Step -1
        col.Remove i
    Next i
    col.Add obj
End Sub

' -----------------------------------------------------------------------------
' Data reload
' -----------------------------------------------------------------------------

Private Sub ReloadStudentsFromDisk()
    Dim errText As String
    Dim students As Collection

    If modStudentsStore.TryLoadStudents(students, errText) Then
        Set m_students = students
    Else
        Dim emptyCol As New Collection
        Set m_students = emptyCol

        ' Show specific reason and allow user to continue editing path etc.
        If Not SuppressLoadErrorPopup Then
            modUiDialogs.ShowError errText, "Порада: натисніть Browse і виберіть коректний students.json."
        End If

    End If

    m_selectedIndex = 0
    ClearFieldsForNew
    RefreshList
End Sub


' -----------------------------------------------------------------------------
' List UI / filtering
' -----------------------------------------------------------------------------

Public Sub chkShowInactive_Click()
    RefreshList
End Sub

Private Sub RefreshList()

    If m_students Is Nothing Then
        Dim emptyCol As New Collection
        Set m_students = emptyCol
    End If
    
    With lstStudents
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "180 pt;0 pt" ' hide id column
    End With

    Dim q As String
    q = Trim$(txtSearch.text)
    If IsPlaceholder(txtSearch) Then q = vbNullString

    Dim st As Dictionary
    For Each st In m_students
        Dim sid As String, nm As String, note As String, chat As String, isActive As Boolean
        sid = NzStr(st("id"))
        nm = NzStr(st("name"))
        chat = NzStr(st("chat_id"))
        note = IIf(st.Exists("note"), NzStr(st("note")), vbNullString)

        isActive = True
        If st.Exists("active") Then isActive = CBool(st("active"))

        If (Not isActive) And (chkShowInactive.value = False) Then
            ' Skip inactive
        Else
            Dim label As String
            label = nm & " (" & sid & ")"
            If Len(note) > 0 Then label = label & " — " & note
            If Not isActive Then label = label & " [inactive]"

            ' Search by name/note/id/chat_id
            Dim searchBlob As String
            searchBlob = nm & " " & note & " " & sid & " " & chat

            If ContainsText(searchBlob, q) Then
                lstStudents.AddItem label
                lstStudents.List(lstStudents.ListCount - 1, COL_ID) = sid
            End If
        End If
    Next st
End Sub

Private Sub lstStudents_Change()
    If lstStudents.ListIndex < 0 Then Exit Sub

    Dim sid As String
    sid = CStr(lstStudents.List(lstStudents.ListIndex, COL_ID))
    LoadToFields sid
End Sub

Private Sub LoadToFields(ByVal sid As String)
    Dim i As Long
    m_selectedIndex = 0

    For i = 1 To m_students.Count
        Dim st As Dictionary
        Set st = m_students(i)

        If NzStr(st("id")) = sid Then
            m_selectedIndex = i

            SetValue txtId, sid
            SetValue txtName, NzStr(st("name"))
            SetValue txtChatId, NzStr(st("chat_id"))
            SetValue txtNote, IIf(st.Exists("note"), NzStr(st("note")), vbNullString)
            chkActive.value = IIf(st.Exists("active"), CBool(st("active")), True)

            Exit Sub
        End If
    Next i
End Sub

Private Sub SelectById(ByVal sid As String)
    Dim i As Long
    For i = 0 To lstStudents.ListCount - 1
        If CStr(lstStudents.List(i, COL_ID)) = sid Then
            lstStudents.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

' -----------------------------------------------------------------------------
' Buttons
' -----------------------------------------------------------------------------

Public Sub btnNew_Click()
    ClearFieldsForNew
End Sub

Public Sub ClearFieldsForNew()
    txtId.text = vbNullString
    txtName.text = vbNullString
    txtChatId.text = vbNullString
    txtNote.text = vbNullString
    chkActive.value = True
    m_selectedIndex = 0

    ApplyPlaceholder txtName
    ApplyPlaceholder txtChatId
    ApplyPlaceholder txtNote
End Sub

Public Sub btnSave_Click()
    Debug.Print "chatValue=[" & txtChatId.text & "]"

    ' Validate + normalize placeholders
    ClearPlaceholder txtName
    ClearPlaceholder txtChatId
    ClearPlaceholder txtNote

    Dim nameValue As String
    nameValue = Trim$(txtName.text)

    Dim chatValue As String
    chatValue = Trim$(txtChatId.text)

    Dim noteValue As String
    noteValue = Trim$(txtNote.text)

    If Len(nameValue) = 0 Then
        MsgBox "Name is required.", vbExclamation, "Validation"
        txtName.SetFocus
        Exit Sub
    End If

    If Len(chatValue) = 0 Then
        MsgBox "Chat ID is required.", vbExclamation, "Validation"
        txtChatId.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(chatValue) Then
        MsgBox "Chat ID must be numeric.", vbExclamation, "Validation"
        txtChatId.SetFocus
        Exit Sub
    End If

    ' Optional: normalize spaces
    chatValue = Replace$(chatValue, " ", vbNullString)
    
    If Not IsNumeric(chatValue) Then
        MsgBox "Chat ID must be numeric.", vbExclamation, "Validation"
        txtChatId.SetFocus
        Exit Sub
    End If
    
    If m_selectedIndex = 0 Then
    
        ' ADD
        Dim stNew As Dictionary
        Set stNew = New Dictionary

        stNew("id") = NextStudentId(m_students)
        stNew("name") = nameValue
        stNew("chat_id") = chatValue
        stNew("active") = CBool(chkActive.value)
        stNew("note") = noteValue

        m_students.Add stNew
        SaveStudents m_students

        ' Reload (keeps file as the source of truth)
        ReloadStudentsFromDisk
        SelectById CStr(stNew("id"))
    Else
        ' UPDATE
        Dim st As Dictionary
        Set st = m_students(m_selectedIndex)

        st("name") = nameValue
        st("chat_id") = chatValue
        st("active") = CBool(chkActive.value)
        st("note") = noteValue

        SaveStudents m_students

        RefreshList
        SelectById CStr(st("id"))
    End If
End Sub

Public Sub btnDeactivate_Click()
    If m_selectedIndex = 0 Then Exit Sub

    Dim st As Dictionary
    Set st = m_students(m_selectedIndex)

    st("active") = False
    SaveStudents m_students

    RefreshList
    SelectById CStr(st("id"))
End Sub

Public Sub btnDelete_Click()
    If m_selectedIndex = 0 Then Exit Sub

    If MsgBox("Delete this student?", vbYesNo + vbExclamation, "Confirm delete") <> vbYes Then Exit Sub

    RemoveFromCollection m_students, m_selectedIndex
    SaveStudents m_students

    m_selectedIndex = 0
    ClearFieldsForNew
    RefreshList
End Sub

Public Sub btnClose_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' Browse path (students.json)
' -----------------------------------------------------------------------------

Public Sub btnBrowseStudentsJson_Click()
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(Type:=msoFileDialogFilePicker)

    With dlg
        .Title = "Select students.json"
        .AllowMultiSelect = False

        .Filters.Clear
        .Filters.Add "JSON files", "*.json"
        .Filters.Add "All files", "*.*"

        .InitialFileName = LoadStudentsJsonPath()

        If .Show <> -1 Then Exit Sub ' user cancelled

        SaveStudentsJsonPath .SelectedItems(1)
    End With

    Me.txtStudentsJsonPath.text = LoadStudentsJsonPath()

    ' Reload list from the new file immediately
    ReloadStudentsFromDisk

    modUiDialogs.ShowInfo "Saved students.json path and reloaded the list."
End Sub

' -----------------------------------------------------------------------------
' Placeholders and search behavior
' -----------------------------------------------------------------------------

Private Sub txtSearch_Enter()
    ClearPlaceholder txtSearch
End Sub

Private Sub txtSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ApplyPlaceholder txtSearch
    RefreshList
End Sub

Private Sub txtSearch_Change()
    ' Do not filter while placeholder text is displayed.
    If IsPlaceholder(txtSearch) Then Exit Sub
    RefreshList
End Sub

Private Sub txtName_Enter(): ClearPlaceholder txtName: End Sub
Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean): ApplyPlaceholder txtName: End Sub

Private Sub txtChatId_Enter(): ClearPlaceholder txtChatId: End Sub
Private Sub txtChatId_Exit(ByVal Cancel As MSForms.ReturnBoolean): ApplyPlaceholder txtChatId: End Sub

Private Sub txtNote_Enter(): ClearPlaceholder txtNote: End Sub
Private Sub txtNote_Exit(ByVal Cancel As MSForms.ReturnBoolean): ApplyPlaceholder txtNote: End Sub

Private Function IsPlaceholder(ByVal tb As MSForms.TextBox) As Boolean
    IsPlaceholder = (tb.ForeColor = PLACEHOLDER_COLOR)
End Function

Private Sub ApplyPlaceholder(ByVal tb As MSForms.TextBox)
    If Len(tb.text) = 0 Then
        tb.text = tb.Tag
        tb.ForeColor = PLACEHOLDER_COLOR
    End If
End Sub

Private Sub ClearPlaceholder(ByVal tb As MSForms.TextBox)
    If tb.ForeColor = PLACEHOLDER_COLOR Then
        If tb.text = tb.Tag Then tb.text = vbNullString
        tb.ForeColor = vbBlack
    End If
End Sub

Private Sub SetValue(ByVal tb As MSForms.TextBox, ByVal value As String)
    tb.ForeColor = vbBlack
    tb.text = value
End Sub

' -----------------------------------------------------------------------------
' Collection helpers
' -----------------------------------------------------------------------------

Private Sub RemoveFromCollection(ByVal col As Collection, ByVal index1 As Long)
    ' Collection.Remove is 1-based.
    col.Remove index1
End Sub

