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
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CONTROL :: bluDownload

'Download a file in the background!

'With thanks to Karl E. Peterson's article that brought this method to light _
 <visualstudiomagazine.com/articles/2008/03/27/simple-asynchronous-downloads.aspx> _
 and this sample usage application sent to me by Tanner Helland _
 <vbforums.com/showthread.php?733409-VB6-Simple-Async-Download-Ctl-for-multiple-Files>

'Status             Ready to use
'Dependencies       Lib.bas
'Last Updated       09-OCT-13
'Last Update        Fix strange bug with not being able to download more than one file

'TODO: Allow multiple simultaneous downloads (use Collection with ID / URL / File)

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

Private My_URL As String
Private My_FilePath As String

'The download is done in the Internet Temporary Files and the filepath is given, _
 we have to copy it out once done
Private TempPath As String

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'As the file downloads
Event Progress( _
    ByVal StatusCode As AsyncStatusCodeConstants, ByVal Status As String, _
    ByVal BytesDownloaded As Long, ByVal BytesTotal As Long _
)
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
    On Error GoTo Fail
    
    'Did some error occur during download?
    If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData _
    Or AsyncProp.BytesMax = 0 Then
        GoTo Fail
    Else
        If Lib.FileExists(My_FilePath) = True Then Call VBA.Kill(My_FilePath)
        Call VBA.FileSystem.FileCopy(TempPath, My_FilePath)
        RaiseEvent Complete
    End If
    
    Exit Sub

Fail:
    RaiseEvent Failed(AsyncProp.StatusCode, AsyncProp.Status)
    Call Me.Cancel
End Sub

'CONTROL AsyncReadProgress _
 ======================================================================================
Private Sub UserControl_AsyncReadProgress(ByRef AsyncProp As AsyncProperty)
    'Provide an event to track progress
    On Error Resume Next
    RaiseEvent Progress( _
        AsyncProp.StatusCode, AsyncProp.Status, _
        AsyncProp.BytesRead, AsyncProp.BytesMax _
    )
    
    'If the temporary file for download was assigned, remember it
    If AsyncProp.StatusCode = vbAsyncStatusCodeCacheFileNameAvailable Then
        Let TempPath = AsyncProp.Status
    End If
    
    'If an error occurs, abort the download
    If AsyncProp.StatusCode = vbAsyncStatusCodeError Then Call Me.Cancel
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    'You can't resize this control, it just appears as a box
    Let UserControl.Width = 32 * Screen.TwipsPerPixelX
    Let UserControl.Height = 32 * Screen.TwipsPerPixelY
    Let UserControl.imgIcon.Width = UserControl.ScaleWidth
    Let UserControl.imgIcon.Height = UserControl.ScaleHeight
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY StatusCodeText : Get a text description for the status codes _
 ======================================================================================
Public Property Get StatusCodeText(ByVal Index As AsyncStatusCodeConstants) As String
    Select Case Index
        'An error occurred during the asynchronous download
        Case vbAsyncStatusCodeError
            Let StatusCodeText = "Error"
        'Finding the resource specified (i.e. DNS lookup)
        Case vbAsyncStatusCodeFindingResource
            Let StatusCodeText = "Finding Resource"
        'Connecting to the resource (i.e. opening TCP/IP)
        Case vbAsyncStatusCodeConnecting
            Let StatusCodeText = "Connecting"
        'Redirection has occured (i.e. HTTP-301/302)
        Case vbAsyncStatusCodeRedirecting
            Let StatusCodeText = "Redirecting"
        'First data received / More data received
        Case vbAsyncStatusCodeBeginDownloadData, vbAsyncStatusCodeEndDownloadData
            Let StatusCodeText = "Downloading"
        'Data has finished downloading
        Case vbAsyncStatusCodeEndDownloadData
            Let StatusCodeText = "Download Complete"
        'A cached copy is being read
        Case vbAsyncStatusCodeUsingCachedCopy
            Let StatusCodeText = "Reading Cache"
        'Sending the HTTP-Request
        Case vbAsyncStatusCodeSendingRequest
            Let StatusCodeText = "Sending Request"
        
        Case Else
            Let StatusCodeText = vbNullString
    End Select
End Property

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
