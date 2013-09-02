VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFAF00&
   Caption         =   "Welcome"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15120
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MaSS1VE.bluControlBox cbxSizer 
      Height          =   360
      Left            =   14760
      TabIndex        =   2
      Top             =   8040
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Style           =   1
      Kind            =   3
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWelcome.frx":0000
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
      Height          =   960
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   5385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To begin, click the levels tab to select a level to edit"
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
      Height          =   270
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   5265
   End
End
Attribute VB_Name = "frmWelcome"
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
'FORM :: frmWelcome

Private Sub Form_Resize()
    Call Me.cbxSizer.Move( _
        Me.ScaleWidth - Me.cbxSizer.Width, Me.ScaleHeight - Me.cbxSizer.Height _
    )
End Sub
