Attribute VB_Name = "Lib"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Lib

'A bunch of common functions VB should have had built-in

'Used in converting colours to Hue / Saturation / Lightness
Public Type HSL
    Hue As Long
    Saturation As Long
    Luminance As Long
End Type

'/// API //////////////////////////////////////////////////////////////////////////////

'Launch a file with its associated application _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762153%28v=vs.85%29.aspx>
Public Declare Function shell32_ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
    ByVal hndWindow As Long, _
    ByVal Operation As String, _
    ByVal File As String, _
    ByVal Parameters As String, _
    ByVal Directory As String, _
    ByVal ShowCmd As SW _
) As Long

Public Enum SW
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

'Get the location of a special folder, e.g. "My Documents", "System32" &c. _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762181%28v=vs.85%29.aspx>
Private Declare Function shfolder_SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" ( _
    ByVal hWndOwner As Long, _
    ByVal Folder As CSIDL, _
    ByVal Token As Long, _
    ByVal Flags As SHGFP, _
    ByVal Path As String _
) As HRESULT

'Full list with descriptions here: _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762494%28v=vs.85%29.aspx>
Public Enum CSIDL
    CSIDL_APPDATA = &H1A&       'Application data (roaming), intended for app data
                                 'that should persist with the user between machines
    CSIDL_LOCAL_APPDATA = &H1C& 'Application data specific to the PC (e.g. cache)
    CSIDL_COMMON_APPDATA = &H23 'Application data shared between all users
    
    
    CSIDL_FLAG_CREATE = &H8000& 'OR this with any of the above to create the folder
                                 'if it doesn't exist (e.g. user deleted My Pictures)
End Enum

Private Enum SHGFP
    SHGFP_TYPE_CURRENT = 0      'Retrieve the folder's current path (it may have moved)
    SHGFP_TYPE_DEFAULT = 1      'Get the default path
End Enum

'Get the location of the temporary files folder _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa364992%28v=vs.85%29.aspx>
Private Declare Function kernel32_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal BufferLength As Long, _
    ByVal Buffer As String _
) As Long

'I need to investigate the actual effectiveness of this lot (preventing repaints to _
 reduce flicker). I've fixed flicker during resizing, but there are instances - mostly _
 when switching level, that several things have to repaint close to each other and I'd _
 like to hold off redrawing the window entirely until the whole process is complete
'--------------------------------------------------------------------------------------
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
'Private Const WM_SETREDRAW as Long = &HB
'Private Const RDW_INVALIDATE as Long = &H1
'Private Const RDW_INTERNALPAINT as Long = &H2
Public Const RDW_UPDATENOW As Long = &H100
Public Const RDW_ALLCHILDREN As Long = &H80
'
'Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
'
'Public Function LockRedraw(ByVal hWnd As Long)
'    Call SendMessage(hWnd, WM_SETREDRAW, 0&, 0&)
'End Function
'
'Public Function UnlockRedraw(ByVal hWnd As Long)
'    Dim r As RECT
'    Call SendMessage(hWnd, WM_SETREDRAW, 1, 0&)
'    Call user32_GetClientRect(hWnd, r)
'    'http://www.xtremevbtalk.com/showthread.php?t=189480
'    Call RedrawWindow(hWnd, r, 0&, RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_UPDATENOW Or RDW_ALLCHILDREN)
'    Call InvalidateRect(hWnd, 0&, 0)
'End Function

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'ArrayDimmed : Is an array dimmed? _
 ======================================================================================
'Taken from: https://groups.google.com/forum/?_escaped_fragment_=msg/microsoft.public.vb.general.discussion/3CBPw3nMX2s/zCcaO-hiCI0J#!msg/microsoft.public.vb.general.discussion/3CBPw3nMX2s/zCcaO-hiCI0J
Public Function ArrayDimmed(varArray As Variant) As Boolean
    Dim pSA As Long
    'Make sure an array was passed in:
    If IsArray(varArray) Then
        'Get the pointer out of the Variant:
        Call blu.kernel32_RtlMoveMemory( _
            ptrDestination:=pSA, ptrSource:=ByVal VarPtr(varArray) + 8, Length:=4 _
        )
        If pSA Then
            'Try to get the descriptor:
            Call blu.kernel32_RtlMoveMemory( _
                ptrDestination:=pSA, ptrSource:=ByVal pSA, Length:=4 _
            )
            'Array is initialized only if we got the SAFEARRAY descriptor:
            Let ArrayDimmed = (pSA <> 0)
        End If
    End If
End Function

'GetSpecialFolder : Get the path to a system folder, e.g. AppData _
 ======================================================================================
Public Function GetSpecialFolder(ByVal Folder As CSIDL) As String
    'Return null should this fail
    Let GetSpecialFolder = vbNullString
    
    'Fill a buffer to receive the path
    Dim Result As String
    Let Result = String$(260, " ")
    'Attempt to get the special folder path, creating it if it doesn't exist _
     (e.g. the user deleted the "My Pictures" folder)
    If shfolder_SHGetFolderPath( _
        0&, Folder Or CSIDL_FLAG_CREATE, 0&, SHGFP_TYPE_CURRENT, Result _
    ) = S_OK Then
        'The string will be null-terminated; find the end and trim, _
         also ensure it always ends in a slash (this can be inconsistent)
        Let GetSpecialFolder = Lib.EndSlash(Left$( _
            Result, InStr(1, Result, vbNullChar) - 1 _
        ))
    End If
End Function

'GetTemporaryFile : Get a unique file name in the temporary files folder _
 ======================================================================================
Public Function GetTemporaryFile() As String
    'The Windows `GetTempFileName` API is not reliable, it has a limit of 65'535 files _
     which could be hit if we generate a lot and the user doesn't clear their cache. _
     Instead we'll use a timestamp that should be sufficient enough
        
    'Generate a unique file name
    Let GetTemporaryFile = Lib.GetTemporaryFolder _
        & App.EXEName & "_" _
        & Year(Now) _
        & Right$("0" & Month(Now), 2) _
        & Right$("0" & Day(Now), 2) _
        & Right$("0" & Hour(Now), 2) _
        & Right$("0" & Minute(Now), 2) _
        & Right$("0" & Second(Now), 2) _
        & "_" & Timer _
        & ".tmp"
End Function

'GetTemporaryFolder : Get the path to the temporary files folder _
 ======================================================================================
Public Function GetTemporaryFolder() As String
    'Return null should this fail
    Let GetTemporaryFolder = vbNullString
    
    'Fill a buffer to receive the path
    Dim Result As String
    Let Result = String$(260, " ")
    
    If kernel32_GetTempPath(Len(Result), Result) > 0 Then
        'The string will be null-terminated; find the end and trim, _
         also ensure it always ends in a slash (this can be inconsistent)
        Let GetTemporaryFolder = Lib.EndSlash(Left$( _
            Result, InStr(1, Result, vbNullChar) - 1 _
        ))
    End If
End Function

'PushByte : Expand an array, adding a byte on the end _
 ======================================================================================
Public Function PushByte(ByVal What As Byte, ByRef Where() As Byte) As Long
    If ArrayDimmed(Where) Then
        'We will return the new length
        Let PushByte = UBound(Where) + 1
        ReDim Preserve Where(LBound(Where) To PushByte) As Byte
    Else
        'Array is not dimmed "()", begin at 0
        Let PushByte = 0
        ReDim Where(0) As Byte
    End If
    'Add the data
    Let Where(PushByte) = What
End Function

'Push Long : Expand an array, adding a long value on the end _
 ======================================================================================
Public Function PushLong( _
    ByVal What As Long, ByRef Where() As Long, _
    Optional ByVal AllowDuplicates As Boolean = True _
) As Boolean
    If ArrayDimmed(Where) Then
        'Is this a duplicate value?
        If AllowDuplicates = False Then
            'Don't add the value; return false to notify this
            If Lib.InArray(What, Where) = True Then Exit Function
        End If
        'Expand the array
        ReDim Preserve Where(LBound(Where) To UBound(Where) + 1) As Long
    Else
        'Array is not dimmed "()", begin at 0
        ReDim LongArray(0) As Long
    End If
    'Add the data
    Let Where(UBound(Where)) = What
    Let PushLong = True
End Function

'BytesToHex : Dump a byte array as a hexadecimal string _
 ======================================================================================
Public Function BytesToHex(var() As Byte) As String
    Dim i As Long
    For i = LBound(var) To UBound(var)
        BytesToHex = BytesToHex & Right$("0" & Hex$(var(i)), 2)
    Next
End Function

'Exists : Check if an item exists in a Collection object _
 ======================================================================================
'<stackoverflow.com/questions/40651/check-if-a-record-exists-in-a-vb6-collection/9535221#9535221>
Public Function Exists(ByVal Key As String, ByRef Col As Collection) As Boolean
    Dim var As Variant
    
TryObject:
    On Error GoTo ExistsTryObject
    Set var = Col(Key)
    Let Exists = True
    Exit Function

TryNonObject:
    On Error GoTo ExistsTryNonObject
    Let var = Col(Key)
    Let Exists = True
    Exit Function

ExistsTryObject:
    'This will reset your Err Handler
    Resume TryNonObject

ExistsTryNonObject:
    Let Exists = False
End Function

'In Array : Check if a long value exists in an array _
 ======================================================================================
Public Function InArray(ByVal What As Long, ByRef Where() As Long) As Boolean
    Let InArray = False
    If Lib.ArrayDimmed(Where) = False Then Exit Function
    
    'This is slow, but it doesn't rely on creating an array error which would crash _
     the executable if we disable array bound checking for speed purposes
    Dim Index As Long
    For Index = LBound(Where) To UBound(Where)
        If Where(Index) = What Then Let InArray = True: Exit Function
    Next
End Function

'CombSort : Sorty an array _
 ======================================================================================
'<www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-%28sort-array-sorting-arrays%29&p=2909248#post2909248>
Public Sub CombSort(ByRef pvarArray As Variant)
    Const ShrinkFactor As Single = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
   
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If pvarArray(i) > pvarArray(i + lngGap) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + lngGap)
                pvarArray(i + lngGap) = varSwap
                blnSwapped = True
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

'RoundUp : Always round a number upwards _
 ======================================================================================
Public Function RoundUp(ByVal InputNumber As Double) As Double
    If Int(InputNumber) = InputNumber _
        Then Let RoundUp = InputNumber _
        Else Let RoundUp = Int(InputNumber) + 1
End Function

'Min : Limit a number to a minimum value _
 ======================================================================================
Public Function Min(ByVal InputNumber As Long, Optional ByVal Minimum As Long = 0) As Long
    Let Min = IIf(InputNumber < Minimum, Minimum, InputNumber)
End Function

'Max : Limit a number to a maximum value _
 ======================================================================================
Public Function Max(ByVal InputNumber As Long, Optional ByVal Maximum As Long = 2147483647) As Long
    Let Max = IIf(InputNumber > Maximum, Maximum, InputNumber)
End Function

'Range : Limit a number to a minimum and maximum value _
 ======================================================================================
Public Function Range( _
    ByVal InputNumber As Long, _
    Optional ByVal Maximum As Long = 2147483647, _
    Optional ByVal Minimum As Long = -2147483648# _
) As Long
    Let Range = Max(Min(InputNumber, Minimum), Maximum)
End Function

'NotZero : Ensure a number is not zero. Useful when dividing by an unknown factor _
 ======================================================================================
Public Function NotZero(ByVal InputNumber As Long, Optional ByVal AtLeast As Long = 1) As Long
    'This is different from Min / Max because it allows you to handle +/- numbers
    If InputNumber = 0 Then Let NotZero = AtLeast Else Let NotZero = InputNumber
End Function

'RGBToHSL : Convert Red, Green, Blue to Hue, Saturation, Lightness _
 ======================================================================================
'<www.xbeat.net/vbspeed/c_RGBToHSL.htm>
Public Function RGBToHSL(ByVal RGBValue As Long) As HSL
    'by Paul - wpsjr1@syix.com, 20011120
    Dim r As Long, G As Long, B As Long
    Dim lMax As Long, lMin As Long
    Dim q As Single
    Dim lDifference As Long
    Static Lum(255) As Long
    Static QTab(255) As Single
    Static init As Long
    
    If init = 0 Then
        For init = 2 To 255 ' 0 and 1 are both 0
            Lum(init) = init * 100 / 255
        Next
        For init = 1 To 255
            QTab(init) = 60 / init
        Next
    End If
    
    r = RGBValue And &HFF
    G = (RGBValue And &HFF00&) \ &H100&
    B = (RGBValue And &HFF0000) \ &H10000
    
    If r > G Then
        lMax = r: lMin = G
    Else
        lMax = G: lMin = r
    End If
    If B > lMax Then
        lMax = B
    ElseIf B < lMin Then
        lMin = B
    End If
    
    RGBToHSL.Luminance = Lum(lMax)
    
    lDifference = lMax - lMin
    If lDifference Then
        'Do a 65K 2D lookup table here for more speed if needed
        RGBToHSL.Saturation = (lDifference) * 100 / lMax
        q = QTab(lDifference)
        Select Case lMax
            Case r
                If B > G Then
                    RGBToHSL.Hue = q * (G - B) + 360
                Else
                    RGBToHSL.Hue = q * (G - B)
                End If
            Case G
                RGBToHSL.Hue = q * (B - r) + 120
            Case B
                RGBToHSL.Hue = q * (r - G) + 240
        End Select
    End If
End Function

'HSLToRGB : Convert Hue, Saturation, Ligthness to (roughly) Red, Green, Blue _
 ======================================================================================
'<www.xbeat.net/vbspeed/c_HSLToRGB.htm>
Public Function HSLToRGB( _
    ByVal Hue As Long, _
    ByVal Saturation As Long, _
    ByVal Luminance As Long _
) As Long
    'by Donald (Sterex 1996), donald@xbeat.net, 20011124
    Dim r As Long, G As Long, B As Long
    Dim lMax As Long, lMid As Long, lMin As Long
    Dim q As Single

    lMax = (Luminance * 255) / 100
  
    If Saturation > 0 Then

        lMin = (100 - Saturation) * lMax / 100
        q = (lMax - lMin) / 60
        
        Select Case Hue
            Case 0 To 60
                lMid = (Hue - 0) * q + lMin
                r = lMax: G = lMid: B = lMin
            Case 60 To 120
                lMid = -(Hue - 120) * q + lMin
                r = lMid: G = lMax: B = lMin
            Case 120 To 180
                lMid = (Hue - 120) * q + lMin
                r = lMin: G = lMax: B = lMid
            Case 180 To 240
                lMid = -(Hue - 240) * q + lMin
                r = lMin: G = lMid: B = lMax
            Case 240 To 300
                lMid = (Hue - 240) * q + lMin
                r = lMid: G = lMin: B = lMax
            Case 300 To 360
                lMid = -(Hue - 360) * q + lMin
                r = lMax: G = lMin: B = lMid
        End Select
        HSLToRGB = B * &H10000 + G * &H100& + r
    Else
        HSLToRGB = lMax * &H10101
    End If
End Function

'EndSlash : Make sure a path always ends in a slash _
 ======================================================================================
Public Function EndSlash(ByVal Path As String) As String
    Let EndSlash = Path
    If Right$(EndSlash, 1) <> "\" Then Let EndSlash = EndSlash & "\"
End Function

'FileExists : See if a file exists or not _
 ======================================================================================
'<cuinl.tripod.com/Tips/fileexist.htm>
Public Function FileExists(ByVal Path As String) As Boolean
    Let FileExists = CBool(Dir$(Path) <> vbNullString)
End Function

'DirExists : See if a folder exists _
 ======================================================================================
'<cuinl.tripod.com/Tips/direxist.htm>
Public Function DirExists(ByVal Path As String) As Boolean
    Let DirExists = CBool(Dir$(Path, vbDirectory) <> vbNullString)
End Function
