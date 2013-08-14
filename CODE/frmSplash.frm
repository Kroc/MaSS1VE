VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFAF00&
   BorderStyle     =   0  'None
   Caption         =   "MaSS1VE"
   ClientHeight    =   5895
   ClientLeft      =   -45
   ClientTop       =   -285
   ClientWidth     =   6945
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5895
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin MaSS1VE.bluControlBox cbxMin 
      Height          =   480
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   1
      Kind            =   1
   End
   Begin MaSS1VE.bluControlBox cbxClose 
      Height          =   480
      Left            =   6480
      TabIndex        =   4
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   1
   End
   Begin MaSS1VE.bluWindow bluWindow 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drag & Drop a Sonic 1 ROM here to begin"
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
      Height          =   210
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1920
      Left            =   2400
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Kroc Camen 2013"
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   2625
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Master System Sonic 1 Visual Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   375
      TabIndex        =   1
      Top             =   840
      Width           =   3225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MaSS1VE"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1635
   End
End
Attribute VB_Name = "frmSplash"
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
'FORM :: frmSplash

Dim ROMVerified As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    Me.BackColor = blu.ActiveColour
End Sub

'FORM OLEDragDrop : A file has been dropped on the form _
 ======================================================================================
Private Sub Form_OLEDragDrop( _
    ByRef Data As DataObject, _
    ByRef Effect As VBRUN.OLEDropEffectConstants, _
    ByRef Button As Integer, ByRef Shift As Integer, _
    ByRef X As Single, ByRef Y As Single _
)
    Call Me.Hide
    ROM.Import (Data.Files(1))
    Call Unload(Me)
End Sub

'FORM OLEDragOver : A file is being dragged over the form _
 ======================================================================================
Private Sub Form_OLEDragOver( _
    ByRef Data As DataObject, _
    ByRef Effect As VBRUN.OLEDropEffectConstants, _
    ByRef Button As Integer, ByRef Shift As Integer, _
    ByRef X As Single, ByRef Y As Single, _
    ByRef State As Integer _
)
    'When something is dragged onto the form, check it before they drop it so that _
     we can show the right pointer
    
    'Don't check with every mouse move, just when it enters for the first time
    If State = VBRUN.DragOverConstants.vbEnter Then
        'Is there file(s) being dragged in?
        If Data.GetFormat(VBRUN.ClipBoardConstants.vbCFFiles) = True Then
            'Read the first file, we'll ignore any others
            Dim BIN As BinaryFile
            Set BIN = New BinaryFile
            Call BIN.Load(Data.Files(1))
            'Verify the first 512 bytes
            If BIN.CRC(0, 512) = &HF150F769 Then
                Let Effect = Effect And vbDropEffectCopy
                Let ROMVerified = True
                Let Me.lblStatus.Caption = "ROM OK. Drop to begin"
            Else
                Let Effect = vbDropEffectNone
                Let ROMVerified = False
                Let Me.lblStatus.Caption = "Sorry, that's not a Sonic 1 Master System ROM"
            End If
            
            Set BIN = Nothing
        End If
    
    ElseIf State = VBRUN.DragOverConstants.vbLeave Then
        Let ROMVerified = False
        Let Me.lblStatus.Caption = "Drag & Drop a Sonic 1 ROM to begin"
    End If
End Sub
