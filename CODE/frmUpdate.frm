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
      _extentx        =   13996
      _extenty        =   10821
   End
   Begin MaSS1VE.bluButton btnUpdate 
      Height          =   480
      Left            =   6120
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
      _extentx        =   2990
      _extenty        =   847
      caption         =   "Exit & Update"
      style           =   1
   End
   Begin MaSS1VE.bluControlBox cbxMin 
      Height          =   480
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      kind            =   1
   End
   Begin MaSS1VE.bluControlBox cbxClose 
      Height          =   480
      Left            =   7560
      TabIndex        =   0
      Top             =   0
      Width           =   480
      _extentx        =   847
      _extenty        =   847
   End
   Begin MaSS1VE.bluWindow bluWindow 
      Left            =   6360
      Top             =   120
      _extentx        =   847
      _extenty        =   847
   End
End
Attribute VB_Name = "frmUpdate"
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
'FORM :: frmUpdate

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
    'Quit the main application, leaving this form loaded
    Unload mdiMain
    'Launch the installer
    Call WIN32.shell32_ShellExecute( _
        Me.hWnd, _
        vbNullString, Run.AppData & "Update.exe", _
        "/UPDATE /D=" & Left$(Run.Path, Len(Run.Path) - 1), _
        Run.AppData, SW_SHOWNORMAL _
    )
    'Unload this form, where upon the application should quit fully
    Unload Me
End Sub

