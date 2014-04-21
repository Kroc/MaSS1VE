VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: BinaryFile

'Reads a file from disk into a byte array and provides an interface to read and write _
 to it easily, including using integers in little-endian order

'Status             Ready to use
'Dependencies       NONE
'Last Updated       28-APR-14
'Last Update        Removed external dependencies

'The CRC function of this class is from "Calculating CRC32 With VB" by Steve McMahon _
 <www.vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp> which was _
 derived from from the VB.NET CRC32 algorithm by Paul Caton

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub kernel32_RtlMoveMemory_ Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'This is the official polynomial used by CRC32 in PKZip. _
 Often the polynomial is shown reversed (04C11DB7)
Private Const CRC_POLYNOMIAL As Long = &HEDB88320
'Store a look up table for the CRC calculations
Private CRCTable(0 To 255) As Long

'This will be the binary file
Private Data() As Byte
'And remember the opened file path for saving
Private File As String

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Dim i As Long, ii As Long
    
    'Create the CRC look-up table _
     (I will not pretend to understand how this works other than it goes through _
      shifting and flipping bits, make of that what you will)
    Dim Value As Long
    For i = 0 To 255
        Value = i
        For ii = 8 To 1 Step -1
            If (Value And 1) Then
                Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                Let Value = Value Xor CRC_POLYNOMIAL
            Else
                Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next
        Let CRCTable(i) = Value
    Next
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase CRCTable
    Erase Data
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY B : Get a single byte value _
 ======================================================================================
Public Property Get B(ByVal Index As Long) As Byte
    Let B = Data(Index)
End Property

Public Property Let B(ByVal Index As Long, ByVal Value As Byte)
    Let Data(Index) = Value
End Property

'PROPERTY IntLE : Get a two-byte integer in Little-Endian order (low, high) _
 ======================================================================================
Public Property Get IntLE(ByVal Index As Long) As Long
    Let IntLE = BytesToInt(Lo:=Data(Index), Hi:=Data(Index + 1))
End Property

Public Property Let IntLE(ByVal Index As Long, ByVal Value As Long)
    'Convert the integer into low/high bytes
    Dim Bytes() As Byte: Let Bytes = IntToBytes(Value)
    'Apply the bytes in Little Endian order (low, high)
    Let Data(Index) = Bytes(0)
    Let Data(Index + 1) = Bytes(1)
End Property

'PROPERTY IntBE : Get a two-byte integer in Big-Endian order (high, low) _
 ======================================================================================
Public Property Get IntBE(ByVal Index As Long) As Long
    Let IntBE = BytesToInt(Lo:=Data(Index + 1), Hi:=Data(Index))
End Property

Public Property Let IntBE(ByVal Index As Long, ByVal Value As Long)
    'Convert the integer into low/high bytes
    Dim Bytes() As Byte: Let Bytes = IntToBytes(Value)
    'Apply the bytes in Big Endian order (high, low)
    Let Data(Index) = Bytes(1)
    Let Data(Index + 1) = Bytes(0)
End Property

'PROPERTY LOF : Length Of the File _
 ======================================================================================
Public Property Get LOF() As Long
    'If no file has been loaded yet, return 0
    If Lib.ArrayDimmed(Data) = False Then Exit Property
    'Return the length of the binary array
    Let LOF = UBound(Data)
End Property

Public Property Let LOF(ByVal Length As Long)
    'Truncate / expand the file
    ReDim Preserve Data(0 To Length - 1) As Byte
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'GetBArr : Get a partial byte-array from the file _
 ======================================================================================
Public Function GetBArr(ByVal Index As Long, ByVal Length As Long) As Byte()
    Dim Arr() As Byte
    ReDim Arr(0 To Length - 1) As Byte
    
    Call kernel32_RtlMoveMemory_( _
        ptrDestination:=Arr(0), ptrSource:=Data(Index), Length:=Length _
    )
    
    Let GetBArr = Arr
End Function

'SetBArr : Set a part of the file data using a byte array _
 ======================================================================================
Public Sub SetBArr(ByVal Index As Long, ByRef Arr() As Byte)
    Call kernel32_RtlMoveMemory_( _
        ptrDestination:=Data(Index), ptrSource:=Arr(0), Length:=UBound(Arr) _
    )
End Sub

'CRC : Perform a Cyclic-Redundancy-Check on all or part of the file _
 ======================================================================================
'Based upon <vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp>
Public Function CRC( _
    Optional ByVal OffsetBegin As Long = 0, Optional ByVal Length As Long = 0 _
) As Long
    'If no length is provided, process through to the end of the file, also _
     if the length requested is longer than the file itself run through to the end
    If Length = 0 Or Length > Me.LOF Then Let Length = UBound(Data) - OffsetBegin
    
    Dim i As Long
    Dim iLookup As Integer
    
    Let CRC = &HFFFFFFFF
    Let Length = Length - 1
    For i = 0 To Length
        iLookup = (CRC And &HFF) Xor Data(OffsetBegin + i)
        CRC = ((CRC And &HFFFFFF00) \ &H100) And 16777215 ' nasty shr 8 with vb :/
        CRC = CRC Xor CRCTable(iLookup)
    Next
   
   Let CRC = Not (CRC)
End Function

'Load : Load a binary file into the class _
 ======================================================================================
Public Function Load(ByVal FilePath As String) As Boolean
    On Error GoTo Fail
    Dim FileNumber As Integer: FileNumber = FreeFile
    Open FilePath For Binary Access Read Lock Write As #FileNumber
    
    'This class holds the whole file in RAM, don't load huge files with it, _
     that would be rather silly, use regular `Seek` and `Get` for anything 1+MB
    ReDim Data(0 To VBA.LOF(FileNumber) - 1) As Byte
    Get #FileNumber, , Data
    'If the load didn't fail, remember the open file to make `Save` easier
    Let File = FilePath
    Let Load = True
    
    Close #FileNumber
    Exit Function
    
Fail:
    Call MsgBox( _
        "Error in BinaryFile.cls: " & _
        "An unexpected error occurred whilst trying to load " & Chr$(34) & FilePath & _
        Chr$(34) & ", the system says: " & _
        vbCrLf & vbCrLf & _
        "Err #" & Err.Number & ": " & Err.Description, _
        vbCritical Or vbOKOnly, "Unexepected Error" _
    )
    'If the file was left open at the point of error, try to free the handle
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
End Function

'Save : Save a binary file from the class _
 ======================================================================================
Public Function Save(Optional ByVal FilePath As String = vbNullString) As Boolean
    'If no filepath is given, use the path of the loaded file
    If FilePath = vbNullString Then Let FilePath = File
    
    'Is the developer doing something stupid?
    If Me.LOF = 0 Then
        'No file has been loaded
        Call MsgBox( _
            "Error in BinaryFile.cls: " & _
            "You can't `Save` when no data has first been Load`ed!", _
            vbCritical Or vbOKOnly, "Developer Error" _
        )
        Exit Function
    End If
    
    'Now try save to disk
    On Error GoTo Fail
    Dim FileNumber As Integer: FileNumber = FreeFile
    Open FilePath For Binary Access Write Lock Read As #FileNumber
    
    Put #FileNumber, , Data
    Let Save = True
    
    Close #FileNumber
    Exit Function

Fail:
    Call MsgBox( _
        "Error in BinaryFile.cls: " & _
        "An unexpected error occurred whilst saving, the system says: " & _
        vbCrLf & vbCrLf & _
        "Err #" & Err.Number & ": " & Err.Description, _
        vbCritical Or vbOKOnly, "Unexepected Error" _
    )
    'If the file was left open at the point of error, try to free the handle
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'BytesToInt : Convert two seperate bytes into a single number _
 ======================================================================================
Private Function BytesToInt(ByVal Lo As Byte, ByVal Hi As Byte) As Long
    'Combine a couple of bytes into a number (for dealing with little-endiness), _
     with thanks to this discssion: <vbforums.com/showthread.php?616645>
    Let BytesToInt = (Hi * 256&) + Lo
End Function

'IntToBytes : Convert a number into two seperate bytes _
 ======================================================================================
Private Function IntToBytes(ByVal Value As Long) As Byte()
    Dim Bytes(0 To 1) As Byte
    'Slice the first 8 bits
    Let Bytes(0) = (Value And &HFF&)
    'and the second 8 bits and shift them down
    Let Bytes(1) = (Value And &HFF00&) / &H100
    Let IntToBytes = Bytes
End Function
