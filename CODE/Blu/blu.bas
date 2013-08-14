Attribute VB_Name = "blu"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Blu

'Stuff shared between the Blu ActiveX controls

'/// PUBLIC VARS //////////////////////////////////////////////////////////////////////

'When a user control is nested in another user control, the `Ambient.UserMode` _
 property returns the incorrect value of True when the control is being run in _
 "Design Mode" (on the form editor). This would cause the design mode controls _
 to be subclassed and crashes the IDE. To stop this, the variable below will _
 always be False when the controls are running in Design Mode. Set `UserMode` _
 to True in your `Sub Main()` to tell the controls it's okay to subclass. _
 (`Sub Main()` will only be run when you your app runs, not during design time)
Public UserMode As Boolean

'The default measurement (px) to base control layout around. _
 Use `blu.Xpx(blu.Metric)` / `blu.Ypx(blu.Metric)` to get it in Twips
Public Const Metric As Long = 32

'The default colour palette for our controls
Public Const BaseColour As Long = vbWhite
Public Const BaseHoverColour As Long = &HEEEEEE
Public Const TextColour As Long = &H999999
Public Const TextHoverColour As Long = &H666666
Public Const ActiveColour As Long = &HFFAF00
Public Const InertColour As Long = &HFFEABA

'The close control box button is red unlike the others
Public Const CloseHoverColour As Long = &H4343E0
Public Const ClosePressColour As Long = &H5050C7 '&H3D3D99

'Public Enums _
 --------------------------------------------------------------------------------------
'The Blu ActiveX controls use these to define friendly names for some properties

Public Enum bluORIENTATION
    Horizontal = 0
    VerticalUp = 1
    VerticalDown = 2
End Enum

Public Enum bluSTATE
    Inactive = 0
    Active = 1
End Enum

Public Enum bluSTYLE
    Normal = 0
    Invert = 1
End Enum

Public Enum bluTRUNCATE
    NoTruncation = 0
    EndElipsis = 1
    MiddleElipsis = 2
End Enum

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'ApplyColoursToForm : Change the colour scheme of the form controls _
 ======================================================================================
Public Sub ApplyColoursToForm( _
    ByRef TheForm As Object, _
    Optional ByVal BaseColour As OLE_COLOR = BaseColour, _
    Optional ByVal TextColour As OLE_COLOR = TextColour, _
    Optional ByVal ActiveColour As OLE_COLOR = ActiveColour, _
    Optional ByVal InertColour As OLE_COLOR = InertColour _
)
    'Deal with all blu controls automatically
    Dim FormControl As Control
    For Each FormControl In TheForm.Controls
        If (TypeOf FormControl Is bluLabel) _
        Or (TypeOf FormControl Is bluButton) _
        Or (TypeOf FormControl Is bluTab) _
        Then
            With FormControl
                On Error Resume Next
                Let .BaseColour = BaseColour
                Let .TextColour = TextColour
                Let .ActiveColour = ActiveColour
                Let .InertColour = InertColour
            End With
        End If
    Next FormControl
End Sub

'Xpx : Shorthand for a number of horizontal pixels converted to twips _
 ======================================================================================
Public Function Xpx(Optional ByVal px As Long = 1) As Long
    'Yes, we could use `Form.ScaleX (...)` but this doesn't require a form and is _
     shorter to write
    Let Xpx = Screen.TwipsPerPixelX * px
End Function

'Ypx : Shorthand for a number of vertical pixels converted to twips _
 ======================================================================================
Public Function Ypx(Optional ByVal px As Long = 1) As Long
    Let Ypx = Screen.TwipsPerPixelY * px
End Function
