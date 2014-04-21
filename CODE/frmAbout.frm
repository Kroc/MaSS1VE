VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFAF00&
   Caption         =   "About"
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
   Begin MaSS1VE.bluHelpView bluHelpView 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5318
      BackColor       =   16756480
   End
End
Attribute VB_Name = "frmAbout"
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
'FORM :: frmAbout

Private Sub Form_Load()
    Call bluHelpView.Navigate(Run.Path & "Help\About.html")
End Sub

Private Sub Form_Resize()
    'If the form is invisible or minimised then don't bother resizing
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Exit Sub
    'Ensure that the MDI child form always stays maximised when changing windows
    If Me.WindowState <> vbMaximized Then Let Me.WindowState = vbMaximized: Exit Sub
    
    'Fill the form with the webview
    Call Me.bluHelpView.Move( _
        0, 0, Me.ScaleWidth, Me.ScaleHeight _
    )
End Sub
