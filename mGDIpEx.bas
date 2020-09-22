Attribute VB_Name = "mGDIpEx"
'   From great stuff:
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'   by Avery
'
'   GDI+ dll:
'   Platform SDK Redistributable: GDI+ RTM
'   http://www.microsoft.com/downloads/release.asp?releaseid=32738

Option Explicit

Public Enum GpStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'//

Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum

Private Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
    ClassID           As CLSID
    FormatID          As CLSID
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

'-- Encoder Parameter structure
Private Type EncoderParameter
    GUID           As CLSID
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type

'-- Encoder Parameters structure
Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type

'-- Encoder parameter sets
Private Const EncoderCompression      As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const EncoderColorDepth       As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const EncoderScanMethod       As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const EncoderVersion          As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const EncoderRenderMethod     As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const EncoderQuality          As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const EncoderTransformation   As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const EncoderLuminanceTable   As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const EncoderChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const EncoderSaveFlag         As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
Private Const CodecIImageBytes        As String = "{025D1823-6C7D-447B-BBDB-A3CBC3DFA2FC}"

'//

Private Const PixelFormat1bppIndexed As Long = &H30101
Private Const PixelFormat4bppIndexed As Long = &H30402
Private Const PixelFormat8bppIndexed As Long = &H30803
Private Const PixelFormat24bppRGB    As Long = &H21808
Private Const PixelFormat32bppARGB   As Long = &H26200A

Private Enum PaletteFlags
    [PaletteFlagsHasAlpha] = &H1
    [PaletteFlagsGrayScale] = &H2
    [PaletteFlagsHalftone] = &H4
End Enum

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type ColorPalette '(8bpp)
   Flags        As PaletteFlags
   Count        As Long
   Entries(255) As RGBQUAD
End Type

'//

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As GpStatus
Private Declare Function GdipGetImageDecodersSize Lib "gdiplus" (numDecoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numDecoders As Long, ByVal Size As Long, Decoders As Any) As GpStatus

Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, BITMAP As Long) As GpStatus
Private Declare Function GdipSetImagePalette Lib "gdiplus" (ByVal hImage As Long, Palette As ColorPalette) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As String, clsidEncoder As CLSID, encoderParams As Any) As GpStatus

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

'==================================================================================================

Private Function GetEncoderClsID(strMimeType As String, ClassID As CLSID)

  Dim Num As Long, Size As Long, i As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    GetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For i = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(PtrToStrW(ICI(i).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(i).ClassID ' Save the Class ID
            GetEncoderClsID = i      ' Return the index number for success
            Exit For
        End If
    Next
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function GetDecoderClsID(strMimeType As String, ClassID As CLSID)

  Dim Num As Long, Size As Long, i As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    GetDecoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageDecodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageDecoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For i = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(PtrToStrW(ICI(i).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(i).ClassID ' Save the Class ID
            GetDecoderClsID = i      ' Return the index number for success
            Exit For
        End If
    Next
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function DEFINE_GUID(ByVal sGuid As String) As CLSID
'-- Courtesy of: Dana Seaman
'   Helper routine to convert a CLSID(aka GUID) string to a structure
'   Example ImageFormatBMP = {B96B3CAB-0728-11D3-9D7B-0000F81EF32E}
    Call CLSIDFromString(StrPtr(sGuid), DEFINE_GUID)
End Function

'-- From www.mvps.org/vbnet
'   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR

Private Function PtrToStrW(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

Private Function PtrToStrA(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenA(lpsz)

    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        PtrToStrA = sOut
    End If
End Function



' SaveDIB256ToGIF inputs:
'
'  DIB              : Main object (image Width and Height)
'  Palette8bpp      : DIB palette entries
'  Dither8bpp       : DIB palette indexes array
'  TransparentEntry : Transparent entry ([0-255] / -1 none)
'  sFilename        : File path

Public Function SaveDIB256ToGIF( _
                DIB As cDIB32, _
                Palette8bpp As cPalette8bpp, _
                Dither8bpp As cDither8bpp, _
                ByVal TransparentEntry As Integer, _
                ByVal sFilename As String _
                ) As Boolean
                
  Dim gplRet As Long
  
  Dim hIm         As Long
  Dim lStride     As Long
  Dim lScan0      As Long
  Dim GdipPalette As ColorPalette
  Dim i           As Long
  
  Dim uEncCLSID   As CLSID
  Dim uEncParams  As EncoderParameters
    
    '-- Build a GDI+ 8bpp bitmap from palette indexes stored in cDither8bpp object
    lStride = ((DIB.Width * 8 + 31) \ 32) * 4 ' Image size
    lScan0 = Dither8bpp.PalIDPtr              ' Image bits (palette indexes)
    gplRet = GdipCreateBitmapFromScan0(DIB.Width, DIB.Height, lStride, [PixelFormat8bppIndexed], ByVal lScan0, hIm)
    
    '-- Build GDI+ palette from palette entries stored in cPalette8bpp object
    With GdipPalette
        .Count = Palette8bpp.MaxCount
        .Flags = [PaletteFlagsHasAlpha]
         CopyMemory .Entries(0), ByVal Palette8bpp.PalettePtr, Palette8bpp.MaxCount * 4
    End With
    
    '-- Set transparent entry
    For i = 0 To Palette8bpp.MaxCount - 1
        '-- Solid color
        GdipPalette.Entries(i).A = 255
    Next i
    Select Case TransparentEntry
      Case 0 To 255
        '-- Set Alpha = 0
        GdipPalette.Entries(TransparentEntry).A = 0
      Case Else
        '-- Ignore
    End Select
    
    '-- Assign palette to bitmap
    gplRet = GdipSetImagePalette(hIm, GdipPalette)
    
    '-- Kill previous
    On Error Resume Next
       Kill sFilename
    On Error GoTo 0
    
    '-- GIF encoder
    GetEncoderClsID "image/gif", uEncCLSID
    '-- Encode
    gplRet = GdipSaveImageToFile(hIm, StrConv(sFilename, vbUnicode), uEncCLSID, uEncParams)

    '-- Free image
    gplRet = GdipDisposeImage(hIm)
    
    '-- Success
    SaveDIB256ToGIF = (gplRet = [OK])
End Function
