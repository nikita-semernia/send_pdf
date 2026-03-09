Attribute VB_Name = "VFMFactory"
Option Explicit
Private Const msMODULE As String = "VFMFactory"

Public Function CreateCLabelControl(ByVal hostForm As Object, _
                                    ByRef ctlsUserFormControls As MSForms.Controls, _
                                    ByRef ctlLabelControl As MSForms.control, _
                                    Optional ByVal boolDefault As Boolean) As CLabelControl
    Dim o As CLabelControl
    Set o = New CLabelControl

    o.SetHostForm hostForm
    o.InitiateProperties ctlsUserFormControls, ctlLabelControl, boolDefault

    Set CreateCLabelControl = o
End Function

Public Function CreateCLabelControlResponder(ByVal oLabelControl As CLabelControl, _
                                             ByRef oLabelControls As CLabelControls) As CLabelControlResponder
    Dim o As CLabelControlResponder
    Set o = New CLabelControlResponder
    o.InitiateProperties oLabelControl, oLabelControls
    Set CreateCLabelControlResponder = o
End Function

Public Function CreateCLabelControlFrameResponder(ByRef ctlFrameControl As control, _
                                                  ByRef oLabelControls As CLabelControls) As CLabelControlFrameResponder
    Dim o As CLabelControlFrameResponder
    Set o = New CLabelControlFrameResponder
    o.InitiateProperties ctlFrameControl, oLabelControls
    Set CreateCLabelControlFrameResponder = o
End Function

Public Function CreateCKeyDownResponder(ByVal hostForm As Object, _
                                        ByRef ctlControl As control, _
                                        ByRef oLabelControls As CLabelControls, _
                                        ByRef ctlsControls As Controls) As CKeyDownResponder
    Dim o As CKeyDownResponder
    Set o = New CKeyDownResponder
    o.InitiateProperties hostForm, ctlControl, oLabelControls, ctlsControls
    Set CreateCKeyDownResponder = o
End Function

Public Function CreateCLabelControls(ByVal hostForm As Object, _
                                     ByRef ctlsControls As MSForms.Controls, _
                                     ByRef arrLabelControlsOrder() As String) As CLabelControls
    Dim o As CLabelControls
    Set o = New CLabelControls
    o.InitiateProperties hostForm, ctlsControls, arrLabelControlsOrder
    Set CreateCLabelControls = o
End Function

Public Function CreateCLabelControlsManager(ByVal hostForm As Object, _
                                            ByRef ctlsControls As MSForms.Controls, _
                                            ByRef arrLabelControlsOrder() As String) As CLabelControlsManager
    Dim o As CLabelControlsManager
    Set o = New CLabelControlsManager
    o.InitiateProperties hostForm, ctlsControls, arrLabelControlsOrder
    Set CreateCLabelControlsManager = o
End Function


