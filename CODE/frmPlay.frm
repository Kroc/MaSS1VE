VERSION 5.00
Begin VB.Form frmPlay 
   BackColor       =   &H00FFAF00&
   Caption         =   "Play"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   9510
   WindowState     =   2  'Maximized
   Begin VB.Frame fraROM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFAF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "sonic the hedgehog.sms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   3120
         UseMnemonic     =   0   'False
         Width           =   2580
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1920
         Left            =   1800
         Picture         =   "frmPlay.frx":0000
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label lblWhatToDo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drag and drop the cartridge below to a folder to save, or double-click it to launch it in an emulator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5580
      End
   End
   Begin VB.Label lblGenerating 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Generating Master System ROM..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2490
      TabIndex        =   3
      Top             =   840
      Width           =   5835
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'FORM :: frmPlay

'ROM export screen

'http://www.vbforums.com/showthread.php?629147-RESOLVED-Drag-and-Drop-vbCFFiles-know-the-destination-folder

'Private Sub picOLEROM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        Call picOLEROM.OLEDrag
'    End If
'End Sub
'
'Private Sub picOLEROM_OLECompleteDrag(Effect As Long)
'    Debug.Print "Dropped"
'End Sub
'
'Private Sub picOLEROM_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'    Let Effect = VBRUN.OLEDropEffectConstants.vbDropEffectCopy
'End Sub
'
'Private Sub picOLEROM_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
'    Let AllowedEffects = VBRUN.OLEDropEffectConstants.vbDropEffectCopy
'    Call Data.Files.Add(Run.Path & "Sonic the Hedgehog (1991)(Sega).sms")
'    Call Data.SetData(, VBRUN.ClipBoardConstants.vbCFFiles)
''    Call Data.SetData( _
''        Run.Path & "Sonic the Hedgehog (1991)(Sega).sms", _
''        VBRUN.ClipBoardConstants.vbCFFiles _
''    )
'End Sub

Private Sub Form_Activate()
    Call ROM.Export( _
        Run.Path & "Sonic_the_hedgehog.sms", _
        Run.Path & "Sonic the Hedgehog (1991)(Sega).sms" _
    )
    Let lblGenerating.Visible = False
    Let fraROM.Visible = True
End Sub

Private Sub Form_Resize()
    Let Me.lblGenerating.Left = (Me.ScaleWidth - Me.lblGenerating.Width) \ 2
'    Call Me.lblGenerating.Move( _
'        (Me.ScaleWidth - Me.lblGenerating.Width) \ 2, _
'        (Me.ScaleHeight - Me.lblGenerating.Height) \ 2 _
'    )
    'Centre the cartridge icon and message, though slightly up looks better
    Call Me.fraROM.Move( _
        (Me.ScaleWidth - Me.fraROM.Width) \ 2, _
        ((Me.ScaleHeight - Me.fraROM.Height) \ 2) - lblWhatToDo.Height _
    )
End Sub
