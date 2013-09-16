VERSION 5.00
Begin VB.UserControl bluDownload 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFAF00&
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "bluDownload.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Picture         =   "bluDownload.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "bluDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CONTROL :: bluDownload

'Download a file in the background!

'With thanks to Karl E. Peterson's article that brought this method to light _
 <visualstudiomagazine.com/articles/2008/03/27/simple-asynchronous-downloads.aspx> _
 and this sample usage application sent to me by Tanner Helland _
 <vbforums.com/showthread.php?733409-VB6-Simple-Async-Download-Ctl-for-multiple-Files>

'Status             INCOMPLETE
'Dependencies       None
'Last Updated       16-SEP-13
'Last Update        I made this

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

Private My_URL As String
Private My_FilePath As String

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'As the file downloads
Event Progress(ByVal BytesDownloaded As Long, ByVal BytesTotal As Long)
'All went well
Event Complete()
'Something went wrong
Event Failed(ByVal StatusCode As AsyncStatusCodeConstants, ByVal Status As String)

'CONTROL AsyncReadComplete _
 ======================================================================================
Private Sub UserControl_AsyncReadComplete(ByRef AsyncProp As AsyncProperty)
    '"Error handling code should be placed in the AsyncReadComplete event procedure, _
      because an error condition may have stopped the download. If this was the case, _
      that error will be raised when the Value property of the AsyncProperty object _
      is accessed." <msdn.microsoft.com/en-us/library/aa445408.aspx>
    On Error GoTo Leave
    
    'Clear the current download so you can start another
    Let My_URL = vbNullString
    Let My_FilePath = vbNullString
    
    'Did some error occur during download?
    If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData _
    Or AsyncProp.BytesMax = 0 Then
        RaiseEvent Failed(AsyncProp.StatusCode, AsyncProp.Status)
        Call Me.Cancel
    Else
        RaiseEvent Complete
    End If
Leave:
End Sub

'CONTROL AsyncReadProgress _
 ======================================================================================
Private Sub UserControl_AsyncReadProgress(ByRef AsyncProp As AsyncProperty)
    'Provide an event to track progress
    On Error Resume Next
    RaiseEvent Progress(AsyncProp.BytesRead, AsyncProp.BytesMax)
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    'You can't resize this control, it just appears as a box
    Let UserControl.Width = 32 * Screen.TwipsPerPixelX
    Let UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Cancel _
 ======================================================================================
Public Function Cancel() As Long
    If My_FilePath = vbNullString Then Exit Function
    
    On Error Resume Next
    Call UserControl.CancelAsyncRead(My_FilePath)
    Let Cancel = Err.Number
    Let My_FilePath = vbNullString
    Let My_URL = vbNullString
End Function

'Download _
 ======================================================================================
Public Function Download( _
    ByVal URL As String, _
    ByVal FilePath As String, _
    Optional ByVal AsyncMode As AsyncReadConstants = vbAsyncReadResynchronize _
) As Long
    'Cancel any existing download before starting the next
    Call Me.Cancel
    
    'Remember the values
    Let My_URL = URL
    Let My_FilePath = FilePath
    
    'Begin downloading
    On Error Resume Next
    Call UserControl.AsyncRead(My_URL, vbAsyncTypeFile, My_FilePath, AsyncMode)
    Let Download = Err.Number
End Function
