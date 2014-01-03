VERSION 5.00
Begin VB.Form frmUpdate 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Update"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MaSS1VE.bluWebView bluWebView 
      Height          =   5775
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10821
   End
   Begin MaSS1VE.bluButton btnUpdate 
      Height          =   480
      Left            =   6120
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   847
      Caption         =   "Exit & Update"
      Style           =   1
   End
   Begin MaSS1VE.bluControlBox cbxMin 
      Height          =   480
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Kind            =   1
   End
   Begin MaSS1VE.bluControlBox cbxClose 
      Height          =   480
      Left            =   7560
      TabIndex        =   0
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MaSS1VE.bluWindow bluWindow 
      Left            =   6360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'FORM :: frmUpdate

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    Let Run.UpdateResponse = vbCancel
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    Call Me.bluWebView.Move( _
        0, Me.cbxClose.Height, Me.ScaleWidth, _
        Me.ScaleHeight - Me.cbxClose.Height - Me.btnUpdate.Height - blu.Ypx(blu.Metric) _
    )
End Sub

'EVENT btnUpdate CLICK _
 ======================================================================================
Private Sub btnUpdate_Click()
    Let Run.UpdateResponse = vbOK
    Unload Me
End Sub

