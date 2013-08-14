VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1Palette"
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
'CLASS :: S1Palette

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public ID As String
Public Title As String

Private Palette(0 To 15) As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Palette
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY Colours _
 ======================================================================================
Public Property Get Colours() As Long(): Let Colours = Palette: End Property
Public Property Let Colours(ByRef NewPalette() As Long)
    Dim i As Long: For i = 0 To 15: Let Palette(i) = NewPalette(i): Next i
End Property

'PROPERTY Colour _
 ======================================================================================
Public Property Get Colour(ByVal Index As Long) As Long: Let Colour = Palette(Index): End Property
Public Property Let Colour(ByVal Index As Long, ByVal NewColour As Long)
    If Index < 0 Or Index > 31 Then Err.Raise (9)
    Let Palette(Index) = NewColour
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'ApplyToImage _
 ======================================================================================
Public Sub ApplyToImage(ByRef DestImage As bluImage)
    'To save a lot of if/thenning in S1Sprites as we apply palette changes, we can _
     ignore images that have not been initialised yet
    If DestImage Is Nothing Then Exit Sub
    'Copy the colours across
    Dim i As Long: For i = 0 To 15: Let DestImage.Colour(i) = Palette(i): Next
End Sub
