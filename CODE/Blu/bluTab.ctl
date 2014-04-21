VERSION 5.00
Begin VB.UserControl bluTab 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   ForeColor       =   &H00FFAF00&
   ScaleHeight     =   495
   ScaleWidth      =   4170
   ToolboxBitmap   =   "bluTab.ctx":0000
   Begin MaSS1VE.bluButton bluButton 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "bluButton"
      State           =   1
   End
   Begin VB.Line lineBorder 
      BorderColor     =   &H00FFAF00&
      X1              =   0
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "bluTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CONTROL :: bluTab

'Status             Planned rewrite
'Dependencies       blu.bas, bluButton.ctl (bluLabel.ctl)
'Last Updated       24-JUL-13

'NOTE: I plan to rewrite this to not depend upon nested controls, it's too unstable _
 and causes flicker

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Private My_ActiveColour As OLE_COLOR
Private My_AutoSize As Boolean
Private My_BaseColour As OLE_COLOR
Private My_Border As Boolean
Private My_CurrentTab As Integer
Private My_Orientation As bluORIENTATION
Private My_TabCount As Integer

Private Const MyTabLength As Long = 1200

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Event TabChanged(ByVal Index As Integer)

'CONTROL InitProperties _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.BaseColour = blu.BaseColour
    Let Me.ActiveColour = blu.ActiveColour
    Let Me.Style = Normal
    
    Let Me.TabCount = 1
    Let Me.CurrentTab = 0
    Let Me.Caption(0) = "Blu Tab"
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        'Set up the tabs
        'TODO: Captions
        Let Me.TabCount = .ReadProperty(Name:="TabCount", DefaultValue:=1)
        Let Me.CurrentTab = .ReadProperty(Name:="CurrentTab", DefaultValue:=0)
        Let Me.Border = .ReadProperty(Name:="Border", DefaultValue:=True)
        Let Me.Orientation = .ReadProperty(Name:="Orientation", DefaultValue:=bluORIENTATION.Horizontal)
        Let Me.AutoSize = .ReadProperty(Name:="AutoSize", DefaultValue:=False)
        
        'Colours
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
    End With
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    Dim i As Integer
    Select Case My_Orientation
        Case bluORIENTATION.Horizontal
            For i = 0 To UserControl.bluButton.Count - 1
            With UserControl.bluButton(i)
                .Top = 0
                If i = 0 Then
                    .Left = 0
                Else
                    .Left = UserControl.bluButton(i - 1).Left + _
                            UserControl.bluButton(i - 1).Width + blu.Xpx
                End If
                .Width = MyTabLength
                .Height = UserControl.ScaleHeight - IIf(My_Border, blu.Ypx, 0)
            End With
            Next
            
            With UserControl.lineBorder
            .Y1 = UserControl.ScaleHeight - blu.Ypx
            .Y2 = .Y1
            .X1 = 0
            .X2 = UserControl.ScaleWidth
            End With
        
        Case bluORIENTATION.VerticalUp
            For i = 0 To UserControl.bluButton.Count - 1
            With UserControl.bluButton(i)
                .Left = 0
                If i = 0 Then
                    .Top = 0
                Else
                    .Top = UserControl.bluButton(i - 1).Top + _
                           UserControl.bluButton(i - 1).Height + blu.Ypx
                End If
                .Width = UserControl.ScaleWidth - IIf(My_Border, blu.Xpx, 0)
                .Height = MyTabLength
            End With
            Next
            
            With UserControl.lineBorder
            .X1 = UserControl.ScaleWidth - blu.Xpx
            .X2 = .X1
            .Y1 = 0
            .Y2 = UserControl.ScaleHeight
            End With
        
        Case bluORIENTATION.VerticalDown
            For i = 0 To UserControl.bluButton.Count - 1
            With UserControl.bluButton(i)
                .Left = IIf(My_Border, blu.Xpx, 0)
                If i = 0 Then
                    .Top = 0
                Else
                    .Top = UserControl.bluButton(i - 1).Top + _
                           UserControl.bluButton(i - 1).Height + blu.Ypx
                End If
                .Width = UserControl.ScaleWidth - IIf(My_Border, blu.Xpx, 0)
                .Height = MyTabLength
            End With
            Next
            
            With UserControl.lineBorder
            .X1 = 0: .X2 = .X1
            .Y1 = 0: .Y2 = UserControl.ScaleHeight
            End With
    End Select
    
    If My_AutoSize = True Then
        If My_Orientation = bluORIENTATION.Horizontal Then
            UserControl.Width = ((MyTabLength * My_TabCount) + My_TabCount - 1)
        Else
            UserControl.Height = ((MyTabLength * My_TabCount) + My_TabCount - 1)
        End If
    End If
End Sub

'CONTROL Terminate _
 ======================================================================================
Private Sub UserControl_Terminate()
    Dim i As Long
    For i = 2 To UserControl.bluButton.Count
        Unload UserControl.bluButton(i - 1)
    Next
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(Name:="AutoSize", Value:=My_AutoSize, DefaultValue:=False)
        Call .WriteProperty(Name:="Border", Value:=My_Border, DefaultValue:=True)
        Call .WriteProperty(Name:="CurrentTab", Value:=My_CurrentTab, DefaultValue:=0)
        Call .WriteProperty(Name:="Orientation", Value:=My_Orientation, DefaultValue:=bluORIENTATION.Horizontal)
        Call .WriteProperty(Name:="TabCount", Value:=My_TabCount, DefaultValue:=1)
        
        Call .WriteProperty(Name:="ActiveColour", Value:=My_ActiveColour, DefaultValue:=blu.ActiveColour)
        Call .WriteProperty(Name:="BaseColour", Value:=My_BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Style", Value:=UserControl.bluButton(0).Style, DefaultValue:=bluSTYLE.Normal)
    End With
End Sub

'EVENT bluButton CLICK : One of the tabs have been clicked! _
 ======================================================================================
Private Sub bluButton_Click(Index As Integer)
    Let Me.CurrentTab = Index
    RaiseEvent TabChanged(Index)
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour: _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR: Let ActiveColour = My_ActiveColour: End Property
Attribute ActiveColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    'Try not to repaint if necessary
    If NewColour = My_ActiveColour Then Exit Property
    Let My_ActiveColour = NewColour
    'Propogate this to each tab button
    Call SetColour
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY AutoSize: Whether to shrink-to-fit the control _
 ======================================================================================
Public Property Get AutoSize() As Boolean: Let AutoSize = My_AutoSize: End Property
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute AutoSize.VB_UserMemId = -500
Public Property Let AutoSize(ByVal Enable As Boolean)
    Let My_AutoSize = Enable
    Call UserControl_Resize
    Call UserControl.PropertyChanged("AutoSize")
End Property

'PROPERTY BaseColour: _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR: Let BaseColour = My_BaseColour: End Property
Attribute BaseColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    'Try not to repaint if necessary
    If NewColour = My_BaseColour Then Exit Property
    Let My_BaseColour = NewColour
    'Propogate this to each tab button
    Call SetColour
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Border: _
 ======================================================================================
Public Property Get Border() As Boolean: Let Border = My_Border: End Property
Attribute Border.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let Border(ByVal Enable As Boolean)
    UserControl.lineBorder.Visible = Enable
    Let My_Border = Enable
    Call UserControl_Resize
    Call UserControl.PropertyChanged("Border")
End Property

'PROPERTY Caption: _
 ======================================================================================
Public Property Get Caption(ByVal Index As Integer) As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    Let Caption = UserControl.bluButton(Index).Caption
End Property

Public Property Let Caption(ByVal Index As Integer, NewCaption As String)
    Let UserControl.bluButton(Index).Caption = NewCaption
    Call UserControl.PropertyChanged("Caption")
End Property

'PROPERTY CurrentTab: _
 ======================================================================================
Public Property Get CurrentTab() As Integer: Let CurrentTab = My_CurrentTab: End Property
Attribute CurrentTab.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute CurrentTab.VB_MemberFlags = "200"
Public Property Let CurrentTab(ByVal Index As Integer)
    Let My_CurrentTab = Index
    
    'Propogate this to each tab button
    Dim i As Integer
    For i = 0 To UserControl.bluButton.Count - 1
        Let UserControl.bluButton(i).State = IIf( _
            Expression:=i = My_CurrentTab, _
            TruePart:=bluSTATE.Active, FalsePart:=bluSTATE.Inactive _
        )
    Next
    
    Call UserControl.PropertyChanged("CurrentTab")
End Property

'PROPERTY Orientation: _
 ======================================================================================
Public Property Get Orientation() As bluORIENTATION: Let Orientation = My_Orientation: End Property
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let Orientation(ByVal NewOrientation As bluORIENTATION)
    'If switching between horizontal / vertical (or vice-versa) then rotate the control
    If ( _
        My_Orientation = bluORIENTATION.Horizontal And _
        (NewOrientation = bluORIENTATION.VerticalDown Or NewOrientation = bluORIENTATION.VerticalUp) And _
        UserControl.Width > UserControl.Height _
    ) Or ( _
        NewOrientation = bluORIENTATION.Horizontal And _
        (My_Orientation = bluORIENTATION.VerticalDown Or My_Orientation = bluORIENTATION.VerticalUp) And _
        UserControl.Height > UserControl.Width _
    ) Then
        Dim Swap As Long
        Let Swap = UserControl.Width
        Let UserControl.Width = UserControl.Height
        Let UserControl.Height = Swap
    End If
    
    Dim bluControl As bluButton
    For Each bluControl In UserControl.bluButton
        bluControl.Orientation = NewOrientation
    Next
    
    Let My_Orientation = NewOrientation
    Call UserControl_Resize
    Call UserControl.PropertyChanged("Orientation")
End Property

'PROPERTY Style: _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = UserControl.bluButton(0).Style: End Property
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    'Change the colours of the tab strip to match the style (normal / invert)
    Call SetColour
    'Propogate this style each tab button
    Dim bluControl As bluButton
    For Each bluControl In UserControl.bluButton
        bluControl.Style = NewStyle
    Next
    
    Call UserControl.PropertyChanged("Style")
End Property

'PROPERTY TabCount: _
 ======================================================================================
Public Property Get TabCount() As Integer: Let TabCount = My_TabCount: End Property
Attribute TabCount.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let TabCount(ByVal Count As Integer)
    If Count <= 0 Then Exit Property
    
    If My_TabCount > 0 Then
        Dim i As Long
        For i = 0 To IIf(Count > My_TabCount, Count - 1, My_TabCount - 1)
            'Is this tab being added or removed?
            If i > My_TabCount - 1 Then
                Load UserControl.bluButton(i)
                With UserControl.bluButton(i)
                    .ActiveColour = Me.ActiveColour
                    .Caption = "Tab " & i + 1
                    .Visible = True
                End With
            ElseIf i > Count - 1 Then
                On Error Resume Next
                Call Unload(UserControl.bluButton(i))
            End If
        Next
    End If
        
    Let My_TabCount = Count
    Let Me.CurrentTab = IIf(Count < My_CurrentTab, Count, My_CurrentTab)
    Call UserControl_Resize
    Call UserControl.PropertyChanged("TabCount")
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'PROPERTY SetColour : Adjust to a colour change and propogate down to the buttons _
 ======================================================================================
Private Sub SetColour()
    'Set the border colour of the tabstrip
    If Me.Style = Invert Then
        Let UserControl.BackColor = My_ActiveColour
        Let UserControl.lineBorder.BorderColor = My_BaseColour
    Else
        Let UserControl.BackColor = My_BaseColour
        Let UserControl.lineBorder.BorderColor = My_ActiveColour
    End If
    'Propogate the colours to each tab button
    Dim bluControl As bluButton
    For Each bluControl In UserControl.bluButton
        Let bluControl.BaseColour = My_BaseColour
        Let bluControl.ActiveColour = My_ActiveColour
    Next
End Sub
