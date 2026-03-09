VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CKeyDownResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE As String = "CKeyDownResponder"

' KeyDown responder for MSForms controls. Its job is to:
' - Allow Enter to trigger the current modernised default button.
' - Allow Tab / arrow keys to cycle through modernised buttons when focus
'   stays on the same control (LastTabbedControl).
' - Reset cycling state on mouse down.

' Supported controls (WithEvents)
Public WithEvents OptionButtonEvents As MSForms.OptionButton
Attribute OptionButtonEvents.VB_VarHelpID = -1
Public WithEvents ScrollBarEvents As MSForms.ScrollBar
Attribute ScrollBarEvents.VB_VarHelpID = -1
Public WithEvents SpinButtonEvents As MSForms.SpinButton
Attribute SpinButtonEvents.VB_VarHelpID = -1
Public WithEvents ListBoxEvents As MSForms.ListBox
Attribute ListBoxEvents.VB_VarHelpID = -1
Public WithEvents TabStripEvents As MSForms.TabStrip
Attribute TabStripEvents.VB_VarHelpID = -1
Public WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1
Public WithEvents TextBoxEvents As MSForms.TextBox
Attribute TextBoxEvents.VB_VarHelpID = -1

Private p_oHostForm As Object
Private p_ctlControl As control
Private p_oLabelControls As CLabelControls
Private p_ctlsControls As Controls

' -----------------------------------------------------------------------------
' Event handlers (forward to shared logic)
' -----------------------------------------------------------------------------

Private Sub OptionButtonEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub OptionButtonEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardMouseDown Button, Shift, X, Y
End Sub

Private Sub ScrollBarEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub SpinButtonEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub ListBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub ListBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardMouseDown Button, Shift, X, Y
End Sub

Private Sub TabStripEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub TabStripEvents_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardMouseDown Button, Shift, X, Y
End Sub

Private Sub CheckBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub CheckBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardMouseDown Button, Shift, X, Y
End Sub

Private Sub TextBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ForwardKeyDown KeyCode, Shift
End Sub

Private Sub TextBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardMouseDown Button, Shift, X, Y
End Sub

Private Sub ComboBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardMouseDown Button, Shift, X, Y
End Sub

' -----------------------------------------------------------------------------
' Shared logic
' -----------------------------------------------------------------------------

Private Sub ForwardKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Const sSOURCE As String = "ForwardKeyDown"
    On Error GoTo ErrorHandler

    HandleKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    End If
End Sub

Private Sub ForwardMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Const sSOURCE As String = "ForwardMouseDown"
    On Error GoTo ErrorHandler

    FormModerniserModule.TabOverflow = 0
    Exit Sub

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    End If
End Sub

Private Sub HandleKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Const sSOURCE As String = "HandleKeyDown"
    On Error GoTo ErrorHandler

    If p_oLabelControls Is Nothing Then Exit Sub
    If p_oLabelControls.LabelControls Is Nothing Then Exit Sub
    If p_oLabelControls.LabelControls.Count = 0 Then Exit Sub

    ' "Default button" is the currently highlighted modern label-control name.
    Dim defaultBtn As String
    defaultBtn = FormModerniserModule.DefaultButton

    ' If the focus stayed on the same control, allow cycling through buttons.
    If p_ctlControl.Name = FormModerniserModule.LastTabbedControl Then

        If KeyCode = vbKeyReturn Then
            TriggerDefaultButton defaultBtn
            Exit Sub
        End If

        If KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
            CycleDefaultButtonForward
            p_oLabelControls.UpdateControlButtonState FormModerniserModule.DefaultButton
            Exit Sub
        End If

        If KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
            CycleDefaultButtonBackward
            p_oLabelControls.UpdateControlButtonState FormModerniserModule.DefaultButton
            Exit Sub
        End If

    Else
        ' Not the last tabbed control: still allow Enter to trigger the default.
        If KeyCode = vbKeyReturn Then
            TriggerDefaultButton defaultBtn
            Exit Sub
        End If

        ' If we had "overflow cycling" active, return focus back.
        If FormModerniserModule.TabOverflow > 0 Then
            On Error Resume Next
            p_ctlsControls(FormModerniserModule.LastTabbedControl).SetFocus
            On Error GoTo 0
        End If
    End If

    Exit Sub

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    End If
End Sub

Private Sub TriggerDefaultButton(ByVal defaultBtn As String)
    If Len(defaultBtn) = 0 Then Exit Sub
    If p_oHostForm Is Nothing Then Exit Sub

    Dim methodName As String
    methodName = sourceButtonName(defaultBtn) & "_Click"

    CallByName p_oHostForm, methodName, VbMethod
End Sub

Private Sub CycleDefaultButtonForward()
    If FormModerniserModule.TabOverflow < p_oLabelControls.LabelControls.Count Then
        FormModerniserModule.TabOverflow = FormModerniserModule.TabOverflow + 1
        FormModerniserModule.DefaultButton = p_oLabelControls.LabelControls.Item(FormModerniserModule.TabOverflow).LabelName
    End If
End Sub

Private Sub CycleDefaultButtonBackward()
    If FormModerniserModule.TabOverflow > 0 Then
        FormModerniserModule.TabOverflow = FormModerniserModule.TabOverflow - 1
        If FormModerniserModule.TabOverflow > 0 Then
            FormModerniserModule.DefaultButton = p_oLabelControls.LabelControls.Item(FormModerniserModule.TabOverflow).LabelName
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' Initialisation
' -----------------------------------------------------------------------------

Public Sub InitiateProperties(ByVal hostForm As Object, _
                              ByRef ctlControl As control, _
                              ByRef oLabelControls As CLabelControls, _
                              ByRef ctlsControls As Controls)

    Const sSOURCE As String = "InitiateProperties"
    On Error GoTo ErrorHandler

    Set p_oHostForm = hostForm
    Set p_ctlControl = ctlControl
    Set p_oLabelControls = oLabelControls
    Set p_ctlsControls = ctlsControls

    Select Case TypeName(ctlControl)
        Case "TextBox"
            Set Me.TextBoxEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name

        Case "OptionButton"
            Set Me.OptionButtonEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name

        Case "ScrollBar"
            Set Me.ScrollBarEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name

        Case "SpinButton"
            Set Me.SpinButtonEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name

        Case "ListBox"
            Set Me.ListBoxEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name

        Case "TabStrip"
            Set Me.TabStripEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name

        Case "CheckBox"
            Set Me.CheckBoxEvents = ctlControl
            FormModerniserModule.LastTabbedControl = ctlControl.Name
    End Select

    Exit Sub

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    End If
End Sub


