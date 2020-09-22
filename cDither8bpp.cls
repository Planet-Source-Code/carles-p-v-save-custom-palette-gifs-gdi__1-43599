VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDither8bpp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' cDither8bpp.cls
' 8bpp dithering [Palette indexes array storing]
' Dependencies: cDIB32.cls, cPalette8bpp.cls
'
' - Color dithering (DitherToColorPalette) original code from VB Accelerator
'   · Pre-reduction of source DIB to 4096 (16x16x16) color space (Speed up)
'   · Weighted RGB channels. See <cDIBPal.CreateOptimal>
'
' - ColourCount function by Robert Rayment

Option Explicit

Public Enum DitherGreyLevels
    [002dgLevels]
    [016dgLevels]
    [256dgLevels]
End Enum

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'==================================================================================================

Public Event Progress(ByVal p As Long)
Public Event ProgressEnd()

'==================================================================================================

Private m_PalID() As Byte ' Palette indexes array

'==================================================================================================

Private Sub Class_Initialize()
    '--
End Sub

Private Sub Class_Terminate()
    Erase m_PalID()
End Sub

'==================================================================================================

Public Function InitializePalID(ByVal nWidth As Long, ByVal nHeight As Long) As Long

    '-- Resize PalID table
    ReDim m_PalID(((nWidth * 8 + 31) \ 32) * 4 - 1, nHeight - 1) As Byte
    
    '-- Return pointer
    InitializePalID = VarPtr(m_PalID(0, 0))
End Function

Public Function PalIDPtr() As Long
    PalIDPtr = VarPtr(m_PalID(0, 0))
End Function

'==================================================================================================

Public Function PalID(ByVal x As Long, ByVal y As Long) As Byte
    PalID = m_PalID(x, y)
End Function


Public Sub DitherToGreyPalette(DIB As cDIB32, Palette As cPalette8bpp, ByVal GreyLevels As DitherGreyLevels, Optional ByVal Diffuse As Boolean = 0)
'-- Floyd-Steinberg error diffusion

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim x As Long, xIn As Long
  Dim y As Long, yIn As Long
  Dim W As Long
  Dim H As Long
  
  Dim tmpEntry() As RGBQUAD
  Dim L As Byte
  
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
    
    If (DIB.hDIB <> 0) Then
    
        pvBuildSA tSA, DIB
        CopyMemory ByVal VarPtrArray(Bits()), VarPtr(tSA), 4

        W = DIB.Width - 1
        H = DIB.Height - 1
        
        '-- Initialize indexes array
        InitializePalID DIB.Width, DIB.Height
        
        '--  Temp. palette entries
        ReDim tmpEntry(255)
        CopyMemory tmpEntry(0), ByVal Palette.PalettePtr, 1024
        
        For y = 0 To H
            For x = 0 To W
                
                '-- Get L
                L = 0.114 * Bits(x, y).B + 0.587 * Bits(x, y).G + 0.299 * Bits(x, y).R
'               L = 0.071 * Bits(x, y).B + 0.707 * Bits(x, y).G + 0.222 * Bits(x, y).R
                
                '-- Save palette index
                Select Case GreyLevels
                  Case [002dgLevels]
                    m_PalID(x, y) = L / 255
                  Case [016dgLevels]
                    m_PalID(x, y) = L / 17
                  Case [256dgLevels]
                    m_PalID(x, y) = L
                End Select
            
                '-- Diffuse error
                If (Diffuse) Then
                
                    errB = CLng(L) - tmpEntry(m_PalID(x, y)).B
                    errG = CLng(L) - tmpEntry(m_PalID(x, y)).G
                    errR = CLng(L) - tmpEntry(m_PalID(x, y)).R
                    Bits(x, y).B = tmpEntry(m_PalID(x, y)).B
                    Bits(x, y).G = tmpEntry(m_PalID(x, y)).G
                    Bits(x, y).R = tmpEntry(m_PalID(x, y)).R
                    
                    '-- Floyd-Steinberg error diffusion...
                    If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                        If (x < W) Then
                            xIn = x + 1
                            newB = Bits(xIn, y).B + (7 * errB) \ 16
                            newG = Bits(xIn, y).G + (7 * errG) \ 16
                            newR = Bits(xIn, y).R + (7 * errR) \ 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            Bits(xIn, y).B = newB
                            Bits(xIn, y).G = newG
                            Bits(xIn, y).R = newR
                        End If
                        If (y < H) Then
                            yIn = y + 1
                            newB = Bits(x, yIn).B + (5 * errB) \ 16
                            newG = Bits(x, yIn).G + (5 * errG) \ 16
                            newR = Bits(x, yIn).R + (5 * errR) \ 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            Bits(x, yIn).B = newB
                            Bits(x, yIn).G = newG
                            Bits(x, yIn).R = newR
                            If (x < W) Then
                                xIn = x + 1
                                newB = Bits(xIn, yIn).B + errB \ 16
                                newG = Bits(xIn, yIn).G + errG \ 16
                                newR = Bits(xIn, yIn).R + errR \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                Bits(xIn, yIn).B = newB
                                Bits(xIn, yIn).G = newG
                                Bits(xIn, yIn).R = newR
                            End If
                            If (x > 0) Then
                                xIn = x - 1
                                newB = Bits(xIn, yIn).B + (3 * errB) \ 16
                                newG = Bits(xIn, yIn).G + (3 * errG) \ 16
                                newR = Bits(xIn, yIn).R + (3 * errR) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                Bits(xIn, yIn).B = newB
                                Bits(xIn, yIn).G = newG
                                Bits(xIn, yIn).R = newR
                            End If
                        End If
                    End If
                  Else
                    Bits(x, y).B = tmpEntry(m_PalID(x, y)).B
                    Bits(x, y).G = tmpEntry(m_PalID(x, y)).G
                    Bits(x, y).R = tmpEntry(m_PalID(x, y)).R
                End If
            Next x
'            RaiseEvent Progress(y)
        Next y
        CopyMemory ByVal VarPtrArray(Bits), 0&, 4
        CopyMemory ByVal VarPtrArray(Bits), 0&, 4
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub DitherToColorPalette(DIB As cDIB32, Palette As cPalette8bpp, Optional ByVal Diffuse As Boolean = 0)
'-- Floyd-Steinberg error diffusion

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim x As Long, xIn As Long
  Dim y As Long, yIn As Long
  Dim W As Long
  Dim H As Long
  
  Dim RGB4096(15, 15, 15) As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim Spd16(255) As Byte
  
  Dim tmpEntry() As RGBQUAD
  Dim hTmpPal    As Long
  Dim nEntry     As Byte
  
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
    
    If (DIB.hDIB <> 0) Then
    
        pvBuildSA tSA, DIB
        CopyMemory ByVal VarPtrArray(Bits()), VarPtr(tSA), 4

        W = DIB.Width - 1
        H = DIB.Height - 1
        
        '-- Initialize indexes array
        InitializePalID DIB.Width, DIB.Height
        
        '-- Speed up for GetNearestPaletteIndex:
        '   Temp. palette entries
            ReDim tmpEntry(255)
            CopyMemory tmpEntry(0), ByVal Palette.PalettePtr, 1024
        '   Temp. logical palette handle
            hTmpPal = Palette.hPal
        
        '-- The real speed up comes here...
        '   This speeds up from 64x64 (aprox.).
        '   We only are reducing available number of colors to 4096, so we only
        '   need to call the expensive GetNearestPaletteIndex function 4096 times.
        
        '-- Build 4096-colors palette indexes table
        For R = 0 To 15
        For G = 0 To 15
        For B = 0 To 15
            RGB4096(R, G, B) = GetNearestPaletteIndex(hTmpPal, RGB(B * 17, G * 17, R * 17))
        Next B, G, R
        
        '-- Speed up translation
        For x = 0 To 255
            Spd16(x) = x / 17
        Next x
        
        For y = 0 To H
            For x = 0 To W
                
                '-- Get pre-calculated palette index:
                nEntry = RGB4096(Spd16(Bits(x, y).R), Spd16(Bits(x, y).G), Spd16(Bits(x, y).B))
                
                '-- Save palette index
                m_PalID(x, y) = nEntry
                                
                '-- Diffuse error
                If (Diffuse) Then
                
                    errB = CLng(Bits(x, y).B) - tmpEntry(nEntry).B
                    errG = CLng(Bits(x, y).G) - tmpEntry(nEntry).G
                    errR = CLng(Bits(x, y).R) - tmpEntry(nEntry).R
                    Bits(x, y) = tmpEntry(nEntry)
                    
                    '-- Floyd-Steinberg error diffusion...
                    If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                        If (x < W) Then
                            xIn = x + 1
                            newB = Bits(xIn, y).B + (7 * errB) \ 16
                            newG = Bits(xIn, y).G + (7 * errG) \ 16
                            newR = Bits(xIn, y).R + (7 * errR) \ 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            Bits(xIn, y).B = newB
                            Bits(xIn, y).G = newG
                            Bits(xIn, y).R = newR
                        End If
                        If (y < H) Then
                            yIn = y + 1
                            newB = Bits(x, yIn).B + (5 * errB) \ 16
                            newG = Bits(x, yIn).G + (5 * errG) \ 16
                            newR = Bits(x, yIn).R + (5 * errR) \ 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            Bits(x, yIn).B = newB
                            Bits(x, yIn).G = newG
                            Bits(x, yIn).R = newR
                            If (x < W) Then
                                xIn = x + 1
                                newB = Bits(xIn, yIn).B + errB \ 16
                                newG = Bits(xIn, yIn).G + errG \ 16
                                newR = Bits(xIn, yIn).R + errR \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                Bits(xIn, yIn).B = newB
                                Bits(xIn, yIn).G = newG
                                Bits(xIn, yIn).R = newR
                            End If
                            If (x > 0) Then
                                xIn = x - 1
                                newB = Bits(xIn, yIn).B + (3 * errB) \ 16
                                newG = Bits(xIn, yIn).G + (3 * errG) \ 16
                                newR = Bits(xIn, yIn).R + (3 * errR) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                Bits(xIn, yIn).B = newB
                                Bits(xIn, yIn).G = newG
                                Bits(xIn, yIn).R = newR
                            End If
                        End If
                    End If
                  Else
                    Bits(x, y) = tmpEntry(nEntry)
                End If
            Next x
'            RaiseEvent Progress(y)
        Next y
        CopyMemory ByVal VarPtrArray(Bits), 0&, 4
        RaiseEvent ProgressEnd
    End If
End Sub

Public Function CountColors(DIB As cDIB32) As Long
'-- by Robert Rayment

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim W As Long, H As Long
  Dim x As Long, y As Long

  Dim BRArr(0 To 255, 0 To 255)  As Integer ' Blue-Red array
  Dim GArr(0 To 255, 0 To 32767) As Byte    ' Green array

  Dim i  As Integer ' Blue
  Dim j  As Integer ' Red
  Dim kG As Byte    ' Green
  Dim k  As Long    ' Green column counter
  Dim kk As Long    ' Column number

    If (DIB.hDIB <> 0) Then

        pvBuildSA tSA, DIB
        CopyMemory ByVal VarPtrArray(Bits()), VarPtr(tSA), 4

        W = DIB.Width - 1
        H = DIB.Height - 1

        For y = 0 To H
            For x = 0 To W

                '-- Get (Blue,Red) coord.
                i = Bits(x, y).B
                j = Bits(x, y).R
                kG = Bits(x, y).G

                If (BRArr(i, j) = 0) Then
                    k = k + 1
                    If (k > 32767) Then ' Rare: give warning message
                        Exit Function   ' or increase GArr size,
                    End If              ' or Redim Preserve which will be slower.

                    '-- Flag BR with column number
                    BRArr(i, j) = k

                    '-- Flag green (0-255) in column k
                    GArr(kG, k) = 1               ' First entry into green column
                    CountColors = CountColors + 1 ' Count unique colors

                  Else
                    kk = BRArr(i, j)
                    '-- Flag green (0-255) in column kk if not flagged
                    If (GArr(kG, kk) = 0) Then
                        GArr(kG, kk) = 1
                        CountColors = CountColors + 1 ' Count unique colors
                    End If
                End If
            Next x
        Next y
        CopyMemory ByVal VarPtrArray(Bits), 0&, 4
    End If
End Function

'==================================================================================================

Private Sub pvBuildSA(tSA As SAFEARRAY2D, DIB As cDIB32)
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.Width
        .pvData = DIB.DIBitsPtr
    End With
End Sub
