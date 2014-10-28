VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluImage"
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
'CLASS :: bluImage

'Create, store and manage an image in memory ("Device Independent Bitmap" -- "DIB"). _
 Supports 8-bit, 24-bit and 32-bit images.

'Status             Ready to use
'Dependencies       blu.bas
'Last Updated       24-JUN-14
'Last Update        Integrated with blu.bas

'By Kroc Camen <kroc@camendesign.com>
'Written using MSDN documentaion with hints and tricks learnt from: _
 * cDIBSection256 by Steve McMahon _
 * iBMP's cDIB by Carles P.V. _
 * MemoryDC, likely by Steve McMahon _
 * pdLayer, part of PhotoDemon, by Tanner Helland

'/// PUBLIC ENUMS /////////////////////////////////////////////////////////////////////
'Enumerations that might be referred to in outside code

Public Enum bluImage_Depth
    [8-Bit] = 8
    [24-Bit] = 24
    [32-Bit] = 32
End Enum

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Private My_hndDC As Long                'Handle to Device Context used for the DIB
Private My_hndDIB As Long               'Handle to the DIB image
Private My_ptrBits As Long              'Pointer to the actual data stream in memory
Private My_Depth As bluImage_Depth      'Bit-depth of the image
Private My_BackgroundColour As Long     'This only applies to 24/32-bit images, _
                                         for 8-bit images the background colour is _
                                         index 0 in the palette
Public IsTransparent As Boolean         'Paint the background transparent?

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'We'll keep hold of the structure used to create the bitmap as this can be used to _
 get and set the data as a byte stream
Private Bitmap As blu.BITMAPINFO

'When we select the DIB into the DC, we need to keep hold of (and restore) the object _
 that might already be in the DC
Private hndOldDIB As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Destroy the DC:
    If My_hndDC <> 0 Then
        'Destroy the DIB:
        If My_hndDIB <> 0 Then
            'Put the object that was already in the DC (when we selected our DIB), _
             back in
            If hndOldDIB <> 0 Then Call gdi32_SelectObject( _
                hndDeviceContext:=My_hndDC, _
                hndGdiObject:=hndOldDIB _
            )
            Call gdi32_DeleteObject(hndGdiObject:=My_hndDIB)
        End If
        Call gdi32_DeleteDC(hndDeviceContext:=My_hndDC)
    End If
    
    'Reset our properties and variables
    Let My_hndDC = 0
    Let My_hndDIB = 0
    Let hndOldDIB = 0
    Let My_ptrBits = 0
    
    'Erase the BITMAPINFO structure
    Call kernel32_RtlZeroMemory(ptrDestination:=Bitmap, Length:=Len(Bitmap))
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'GET BackgroundColour _
 ======================================================================================
Public Property Get BackgroundColour() As Long
    'For 24/32-bit images the background colour is stored in the property, _
     for 8-bit images we refer to index 0 of the palette
    If My_Depth = [8-Bit] _
        Then Let BackgroundColour = Me.Colour(0) _
        Else Let BackgroundColour = My_BackgroundColour
End Property

'GET Colour : Get an individual palette colour: _
 ======================================================================================
Public Property Get Colour(ByVal Index As Byte) As OLE_COLOR
    Dim QuadColour As blu.RGBQUAD
    
    'Try to fetch the colour
    If gdi32_GetDIBColorTable( _
        hndDeviceContext:=My_hndDC, _
        StartIndex:=Index, Count:=1, _
        ptrRGBQUAD:=QuadColour _
    ) = 1 Then
        Let Colour = VBA.RGB(QuadColour.Red, QuadColour.Green, QuadColour.Blue)
    Else
        Let Colour = -1
    End If
End Property

'LET Colour : Set an individual palette colour: _
 ======================================================================================
Public Property Let Colour(ByVal Index As Byte, NewColour As OLE_COLOR)
    'Should the colour be a system colour (like "button face"), get the real colour
    Let NewColour = blu.OleTranslateColor(NewColour)
    
    'Prepre the colour structure to use with the API call
    Dim QuadColour As blu.RGBQUAD
    With QuadColour
        .Blue = (NewColour And &HFF0000) \ &H10000
        .Green = (NewColour And &HFF00&) \ &H100
        .Red = (NewColour And &HFF)
    End With
    
    'Push the colour in
    Call gdi32_SetDIBColorTable( _
        hndDeviceContext:=My_hndDC, _
        StartIndex:=Index, Count:=1, _
        ptrRGBQUAD:=QuadColour _
    )
End Property

'GET Depth : Expose the bit-depth _
 ======================================================================================
Public Property Get Depth() As bluImage_Depth: Let Depth = My_Depth: End Property

'GET hDC : Expose the "hDC" for external drawing operations _
 ======================================================================================
Public Property Get hDC() As Long: Let hDC = My_hndDC: End Property

Public Property Get hDIB() As Long: Let hDIB = My_hndDIB: End Property

'GET Height _
 ======================================================================================
Public Property Get Height() As Long
    'Remember that the height is negative to flip the image the right way up!
    If My_hndDC <> 0 And My_hndDIB <> 0 _
        Then Let Height = Abs(Bitmap.Header.Height) _
        Else Let Height = 0
End Property

'GET Width _
 ======================================================================================
Public Property Get Width() As Long
    If My_hndDC <> 0 And My_hndDIB <> 0 _
        Then Let Width = Bitmap.Header.Width _
        Else Let Width = 0
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Cls : Clear the image, filling it with the background colour _
 ======================================================================================
Public Sub Cls()
    'Define the bounding rectangle to fill
    Dim Box As RECT
    Call user32_SetRect( _
        RECTToSet:=Box, Left:=0, Top:=0, Right:=Me.Width, Bottom:=Me.Height _
    )
    'Set the background colour to fill
    Call gdi32_SetDCBrushColor( _
        My_hndDC, Me.BackgroundColour _
    )
    'Do the filling
    Call user32_FillRect( _
        My_hndDC, Box, _
        blu.gdi32_GetStockObject(DC_BRUSH) _
    )
End Sub

'Create8Bit : Create a new 8-bit paletted image _
 ======================================================================================
Public Function Create8Bit( _
    ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
    Optional ByRef Palette_LongArray As Variant = Null, _
    Optional ByVal UseTransparency As Boolean = False _
) As Long
    'Create the image by proxy
    Let Create8Bit = Create( _
        ImageWidth:=ImageWidth, ImageHeight:=ImageHeight, _
        BitDepth:=[8-Bit] _
    )
    
    'Has a palette been provided?
    If IsMissing(Palette_LongArray) = False And _
       IsNull(Palette_LongArray) = False _
    Then
        On Error Resume Next
        Dim i As Long 'Copy the palette in
        For i = LBound(Palette_LongArray) To UBound(Palette_LongArray)
            Let Me.Colour(i) = Palette_LongArray(i)
        Next
    End If
    
    'If marked as transparent then all painting operations will exclude the background
    Let Me.IsTransparent = UseTransparency
End Function

'Create24Bit : Create a new 24-bit image _
 ======================================================================================
Public Function Create24Bit( _
    ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
    Optional ByVal BackgroundColour As OLE_COLOR = vbWhite, _
    Optional ByVal UseTransparency As Boolean = False _
) As Long
    'Create the image by proxy
    Let Create24Bit = Create( _
        ImageWidth:=ImageWidth, ImageHeight:=ImageHeight, _
        BitDepth:=[24-Bit] _
    )
    
    'If marked as transparent then all painting operations will exclude the background
    Let My_BackgroundColour = BackgroundColour
    Let Me.IsTransparent = UseTransparency
    
    'Paint the background colour on
    Call Me.Cls
End Function

'Create32Bit : Create a new 32-bit image _
 ======================================================================================
'TODO: The plan here is to use a 32-bit transparent background and always _
 `AlphaBlend` in `Paint` calls
Public Function Create32Bit( _
    ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
    Optional ByVal BackgroundColour As OLE_COLOR = vbWhite, _
    Optional ByVal UseTransparency As Boolean = False _
) As Long
    'Create the image by proxy
    Let Create32Bit = Create( _
        ImageWidth:=ImageWidth, ImageHeight:=ImageHeight, _
        BitDepth:=[32-Bit] _
    )
    
    'If marked as transparent then all painting operations will exclude the background
    Let My_BackgroundColour = BackgroundColour
    Let Me.IsTransparent = UseTransparency
    
    'Paint the background colour on (24-bit for now)
    Call Me.Cls
End Function

'GetByteStream : Get the raw data stream as a byte array _
 ======================================================================================
Public Function GetByteStream() As Byte()
    'Create a buffer of the right size
    Dim ByteStream() As Byte
    ReDim ByteStream(0 To Bitmap.Header.DataSize - 1) As Byte
    'Call the API to copy the data into our buffer
    Call gdi32_GetDIBits( _
        hndDeviceContext:=My_hndDC, _
        hndDIB:=My_hndDIB, _
        StartScan:=0, NumberOfScans:=Abs(Bitmap.Header.Height), _
        ptrBits:=ByteStream(0), ptrBITMAPINFO:=Bitmap, Usage:=DIB_RGB_COLORS _
    )
    Let GetByteStream = ByteStream
End Function

'SetByteStream : Set the image pixels using a byte array _
 ======================================================================================
Public Function SetByteStream(ByRef ByteStream() As Byte)
    Call gdi32_SetDIBits( _
        hndDeviceContext:=My_hndDC, _
        hndDIB:=My_hndDIB, _
        StartScan:=0, NumberOfScans:=Abs(Bitmap.Header.Height), _
        ptrBits:=ByteStream(0), ptrBITMAPINFO:=Bitmap, _
        Usage:=DIB_RGB_COLORS _
    )
End Function

'Paint : Draw the image onto something _
 ======================================================================================
'NOTE: If you have to do a lot of painting in a loop, it's best to make the `BitBlt`, _
 `StretchBlt` or `TransparentBlt` calls yourself rather than use this function, _
 it'll be a _lot_ faster
Public Function Paint( _
    ByVal hndDeviceContext As Long, _
    Optional ByVal DestLeft As Long = 0, Optional ByVal DestTop As Long = 0, _
    Optional ByVal DestWidth As Long = 0, Optional ByVal DestHeight As Long = 0, _
    Optional ByVal SrcLeft As Long = 0, Optional ByVal SrcTop As Long = 0, _
    Optional ByVal SrcWidth As Long = 0, Optional ByVal SrcHeight As Long = 0 _
) As Long
    If DestWidth <= 0 Then DestWidth = Me.Width
    If DestHeight <= 0 Then DestHeight = Me.Height
    If SrcWidth <= 0 Then SrcWidth = DestWidth
    If SrcHeight <= 0 Then SrcHeight = DestHeight
    
    'Should we exclude the background? (use transparency)
    If IsTransparent = True Then
        Let Paint = gdi32_GdiTransparentBlt( _
            hndDestDC:=hndDeviceContext, _
            DestLeft:=DestLeft, DestTop:=DestTop, _
            DestWidth:=DestWidth, DestHeight:=DestHeight, _
            hndSrcDC:=My_hndDC, _
            SrcLeft:=SrcLeft, SrcTop:=SrcTop, _
            SrcWidth:=SrcWidth, SrcHeight:=SrcHeight, _
            TransparentColour:=Me.BackgroundColour _
        )
    Else
        'Do we need to stretch?
        If SrcWidth <> DestWidth Or SrcHeight <> DestHeight Then
            Let Paint = gdi32_StretchBlt( _
                hndDestDC:=hndDeviceContext, _
                DestLeft:=DestLeft, DestTop:=DestTop, _
                Width:=DestWidth, Height:=DestHeight, _
                hndSrcDC:=My_hndDC, _
                SrcLeft:=SrcLeft, SrcTop:=SrcTop, _
                SrcWidth:=SrcWidth, SrcHeight:=SrcHeight, _
                RasterOperation:=vbSrcCopy _
            )
        Else
            'Do a simple straight copy without stretching:
            Let Paint = gdi32_BitBlt( _
                hndDestDC:=hndDeviceContext, _
                DestLeft:=DestLeft, DestTop:=DestTop, _
                Width:=DestWidth, Height:=DestHeight, _
                hndSrcDC:=My_hndDC, _
                SrcLeft:=SrcLeft, SrcTop:=SrcTop, _
                RasterOperation:=vbSrcCopy _
            )
        End If
    End If
End Function

'Save : Save image to a bitmap file _
 ======================================================================================
'After seeing many examples I find it odd that this code works and is yet so much _
 simpler and more striaght-forward than anybody else's
'NOTE: For reasons not understood the use of `LenB` here instead of `Len` creates _
 an invalid bitmap file, I was made to think that `LenB` was the safest option
Public Sub Save(ByVal FilePath As String)
    'Prepare a bitmap file header
    Dim FileHeader As BITMAPFILEHEADER
    With FileHeader
        'The bitmap identifier at the beginning of the file
        Let .Type = &H4D42
        'The length of the file (size of the header plus the image data)
        Let .Size = Len(FileHeader) + Bitmap.Header.DataSize
        'How big the header is, i.e. where does the image data begin?
        Let .OffsetToBits = Len(FileHeader) + Len(Bitmap)
    End With
    
    'Delete the image if it's already there
    If FileExists(FilePath) = True Then VBA.Kill (FilePath)
    
    'Write the bitmap to disk
    On Error GoTo Fail
    Dim FileNumber As Integer: Let FileNumber = FreeFile
    Open FilePath For Binary Access Write Lock Read As #FileNumber
        'Headers
        Put #FileNumber, , FileHeader
        Put #FileNumber, , Bitmap
        'Image data
        Put #FileNumber, , Me.GetByteStream
    Close #FileNumber
    Exit Sub
    
Fail:
    Call MsgBox( _
        "Error in bluImage.cls: " & _
        "An unexpected error occurred whilst saving, the system says: " & _
        vbCrLf & vbCrLf & _
        "Err #" & Err.Number & ": " & Err.Description, _
        vbCritical Or vbOKOnly, "Unexepected Error" _
    )
    'If the file was left open at the point of error, try to free the handle
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Create : Create an image _
 ======================================================================================
'This handles the real work for `Create8/24/32Bit` wrapper functions
Private Function Create( _
    ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
    ByVal BitDepth As bluImage_Depth _
) As Long
    'Destroy any existing image
    Call Class_Terminate
    
    'Expose the bit-depth chosen
    Let My_Depth = BitDepth
    
    'Define the bitmap
    With Bitmap.Header
        .SizeOfMe = Len(Bitmap.Header)
        .BitPlanes = 1
        .Width = ImageWidth
        .Height = -ImageHeight      'Flip the image to be top-down, as you expect
        .Depth = BitDepth
        .Compression = BI_RGB
        'Align the rows to 4 bytes when calculating the data size
        .DataSize = (((ImageWidth * BitDepth + 31) \ 32) * 4) * ImageHeight
    End With
    
    If My_Depth = [8-Bit] Then
        'Set a quick and dirty test palette
        Dim i As Long: For i = 0 To 255: With Bitmap.Colors(i)
            .Blue = (i Mod 16) * 16: .Green = (i Mod 16) * 16: .Red = (i Mod 16) * 16
        End With: Next
    End If
    
    'Create a Device Context based on the screen
    Let My_hndDC = gdi32_CreateCompatibleDC(hndDeviceContext:=0)
    If My_hndDC <> 0 Then
        'Proceed to creating the DIB
        Let My_hndDIB = gdi32_CreateDIBSection( _
            hndDeviceContext:=My_hndDC, _
            ptrBITMAPINFO:=Bitmap, _
            Usage:=DIB_RGB_COLORS, _
            ptrBits:=My_ptrBits, _
            hndFileMappingObject:=0, Offset:=0 _
        )
        If My_hndDIB = 0 Then
            'If this failed, clean up and fail
            Call Class_Terminate
        Else
            'Select the DIB into the DC. An object might already be in there, _
             so keep hold of it
            Let hndOldDIB = gdi32_SelectObject( _
                hndDeviceContext:=My_hndDC, _
                hndGdiObject:=My_hndDIB _
            )
        End If
    End If
    
    'Return with success / failure
    Let Create = (My_hndDC <> 0)
End Function

'FileExists : See if a file exists or not _
 ======================================================================================
Public Function FileExists(fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Let Temp = VBA.GetAttr(fName)
    Let FileExists = Not CBool(Err)
End Function
