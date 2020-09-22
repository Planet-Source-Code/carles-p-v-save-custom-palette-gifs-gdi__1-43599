VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' cDIB32.cls (32bpp DIB section)
'
' From cDIBSection.cls
' Copyright © 1999 Steve McMahon

Option Explicit

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    BmBits       As Long
End Type

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

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DIB_RGB_COLORS  As Long = 0
Private Const OBJ_BITMAP      As Long = 7
Private Const SRCCOPY         As Long = &HCC0020
Private Const COLORONCOLOR    As Long = 3
Private Const HALFTONE        As Long = 4
Private Const CF_BITMAP       As Long = 2

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'==================================================================================================

Public Enum DibStretchCts
    [dsColorOnColor] = COLORONCOLOR
    [dsHalftone] = HALFTONE
End Enum

Public Event Progress(ByVal p As Long)
Public Event ProgressEnd()

'==================================================================================================

Private m_tBIH    As BITMAPINFOHEADER
Private m_hDC     As Long
Private m_hDIB    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long

'==================================================================================================

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

'==================================================================================================

Public Function Create(ByVal NewWidth As Long, ByVal NewHeight As Long, Optional ByVal BackColor As Long = vbBlack) As Boolean
    
    '-- Destroy previous
    Destroy
    
    '-- Prepare header
    With m_tBIH
        .biSize = Len(m_tBIH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = NewWidth
        .biHeight = -NewHeight
        .biSizeImage = (4 * NewWidth) * NewHeight
    End With
    
    '-- Create DIB section
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        '-- Create DIB
        m_hDIB = CreateDIBSection(m_hDC, m_tBIH, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If (m_hDIB <> 0) Then
            '-- Select into a DC
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
            '-- Set backColor
            Cls BackColor
          Else
            Destroy
        End If
    End If
    
    '-- Success
    Create = (m_hDIB <> 0)
End Function

Public Function CreateFromStdPicture(Picture As StdPicture) As Byte

  Dim tBI      As BITMAP
  Dim lhDC     As Long
  Dim lhOldBmp As Long
    
    If (Not Picture Is Nothing) Then

        If (GetObjectType(Picture) = OBJ_BITMAP) Then
            GetObject Picture, Len(tBI), tBI
            
            If (Create(tBI.bmWidth, tBI.bmHeight)) Then
                lhDC = CreateCompatibleDC(0)
                If (lhDC <> 0) Then
                    lhOldBmp = SelectObject(lhDC, Picture.handle)
                    
                    '-- Load Bits (-> 32bpp)
                    BitBlt m_hDC, 0, 0, tBI.bmWidth, tBI.bmHeight, lhDC, 0, 0, SRCCOPY
                    
                    '-- Destroy temp. DC
                    SelectObject lhDC, lhOldBmp
                    DeleteDC lhDC
                    
                    '-- Success
                    CreateFromStdPicture = 32
                End If
            End If
        End If
    End If
End Function

Public Sub Cls(Optional ByVal BackColor As Long = &H0)

  Dim tRect  As RECT2
  Dim hBrush As Long
    
    If (Me.hDIB <> 0) Then
        SetRect tRect, 0, 0, m_tBIH.biWidth, -m_tBIH.biHeight
        hBrush = CreateSolidBrush(BackColor)
        FillRect m_hDC, tRect, hBrush
        DeleteObject hBrush
    End If
End Sub

Public Sub LoadBlt(ByVal hSrcDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)
    
    If (Me.hDIB <> 0) Then
        BitBlt m_hDC, 0, 0, m_tBIH.biWidth, -m_tBIH.biHeight, hSrcDC, x, y, SRCCOPY
    End If
End Sub

Public Sub LoadDIBBlt(DIB As cDIB32)
    
    If (Me.hDIB <> 0 And DIB.hDIB <> 0) Then
        DIB.Stretch m_hDC, 0, 0, m_tBIH.biWidth, -m_tBIH.biHeight, 0, 0, DIB.Width, DIB.Height
    End If
End Sub

'==================================================================================================

Public Property Get hDIBDC() As Long
    hDIBDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get DIBitsPtr() As Long
    DIBitsPtr = m_lpBits
End Property

Public Property Get Size() As Long
    Size = m_tBIH.biSizeImage
End Property

Public Property Get BytesPerScanLine() As Long
    BytesPerScanLine = 4 * m_tBIH.biWidth
End Property

Public Property Get Width() As Long
    Width = m_tBIH.biWidth
End Property

Public Property Get Height() As Long
    Height = -m_tBIH.biHeight
End Property

'==================================================================================================

Public Sub Paint(ByVal hDstDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal ScaleFactor As Long = 1)
    
  Dim lRet As Long
  
    If (Me.hDIB <> 0) Then
        With m_tBIH
            lRet = StretchBlt(hDstDC, x, y, .biWidth * ScaleFactor, -.biHeight * ScaleFactor, m_hDC, 0, 0, .biWidth, -.biHeight, SRCCOPY)
        End With
    End If
End Sub

Public Sub Stretch(ByVal hDstDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, Optional ByVal StretchColorMode As DibStretchCts = [dsColorOnColor])
  
  Dim oldMode As Long
  Dim lRet    As Long
    
    If (Me.hDIB <> 0) Then
        oldMode = GetStretchBltMode(hDstDC)
        SetStretchBltMode hDstDC, StretchColorMode
        lRet = StretchBlt(hDstDC, x, y, nWidth, nHeight, m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, SRCCOPY)
        SetStretchBltMode hDstDC, oldMode
    End If
End Sub

Public Sub Destroy()

    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            SelectObject m_hDC, m_hOldDIB
            DeleteObject m_hDIB
        End If
        DeleteDC m_hDC
    End If
    m_hDC = 0: m_hDIB = 0: m_hOldDIB = 0: m_lpBits = 0
End Sub

Public Sub Orientation(ByVal Rotate90 As Boolean, ByVal FlipH As Boolean, ByVal FlipV As Boolean)
  
  Dim sDIB    As New cDIB32
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim x As Long
  Dim y As Long
  Dim newW As Long, W As Long
  Dim newH As Long, H As Long
  
    If (Me.hDIB <> 0) Then
            
        W = Me.Width
        H = Me.Height
        
        If ((FlipH And Not Rotate90) Or (FlipV And Rotate90)) Then
            StretchBlt Me.hDIBDC, W - 1, 0, -W, H, Me.hDIBDC, 0, 0, W, H, SRCCOPY
        End If
        If ((FlipV And Not Rotate90) Or (FlipH And Rotate90)) Then
            StretchBlt Me.hDIBDC, 0, H - 1, W, -H, Me.hDIBDC, 0, 0, W, H, SRCCOPY
        End If
        
        If (Rotate90) Then
        
            '-- Get source Bits
            sDIB.Create Me.Width, Me.Height
            sDIB.LoadBlt Me.hDIBDC
            pvBuildSA stSA, sDIB
            CopyMemory ByVal VarPtrArray(sBits()), VarPtr(stSA), 4
            
            '-- Create new DIB
            Create Me.Height, Me.Width
            pvBuildSA dtSA, Me
            CopyMemory ByVal VarPtrArray(dBits()), VarPtr(dtSA), 4
        
            newW = Me.Height - 1
            newH = Me.Width - 1
            
            For y = 0 To newH
                For x = 0 To newW
                    dBits(y, newW - x) = sBits(x, y)
                Next x
'                RaiseEvent Progress(y)
            Next y
            CopyMemory ByVal VarPtrArray(sBits), 0&, 4
            CopyMemory ByVal VarPtrArray(dBits), 0&, 4
        End If
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub GetBestFitInfo(ByVal DstW As Long, ByVal DstH As Long, bfx As Long, bfy As Long, bfW As Long, bfH As Long, Optional ByVal StretchFit As Boolean = 0)
  
  Dim cW As Single
  Dim cH As Single
    
    bfx = 0
    bfy = 0
    If ((Me.Width > DstW Or Me.Height > DstH) Or StretchFit) Then
        cW = DstW / Me.Width
        cH = DstH / Me.Height
        If (cW < cH) Then
            bfW = DstW
            bfH = Me.Height * cW
          Else
            bfH = DstH
            bfW = Me.Width * cH
        End If
      Else
        bfW = Me.Width
        bfH = Me.Height
    End If
    bfx = (DstW - bfW) \ 2
    bfy = (DstH - bfH) \ 2
End Sub

'==================================================================================================

Public Sub CopyToClipboard()
    
  Dim dsk_Wnd     As Long
  Dim dsk_hDC     As Long
  
  Dim clp_hDC     As Long
  Dim clp_hObj    As Long
  Dim clp_hOldObj As Long
    
    If (Me.hDIB <> 0) Then
    
        If (OpenClipboard(0) <> 0) Then
        
            dsk_Wnd = GetDesktopWindow
            dsk_hDC = GetDC(dsk_Wnd)
            
            clp_hDC = CreateCompatibleDC(dsk_hDC)
            If (clp_hDC <> 0) Then
            
                clp_hObj = CreateCompatibleBitmap(dsk_hDC, Me.Width, Me.Height)
                If (clp_hObj <> 0) Then
                    
                    clp_hOldObj = SelectObject(clp_hDC, clp_hObj)
                    Me.Paint clp_hDC
                    SelectObject clp_hDC, clp_hOldObj
        
                    EmptyClipboard
                    SetClipboardData CF_BITMAP, clp_hObj
                    CloseClipboard
        
                    DeleteDC clp_hDC
                End If
            End If
            
            ReleaseDC dsk_Wnd, dsk_hDC
        End If
    End If
End Sub

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

