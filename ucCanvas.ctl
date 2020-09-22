VERSION 5.00
Begin VB.UserControl ucCanvas 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   Begin VB.PictureBox iCanvas 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1185
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   0
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "ucCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ucCanvas.ctl [Removed crop feature and other simplifications]
' DIB Viewer [Scroll/Zoom/Crop]
'
' Carles P.V.

Option Explicit

Public WithEvents DIB As cDIB32
Attribute DIB.VB_VarHelpID = -1

Public Event DIBProgress(ByVal p As Long)
Public Event DIBProgressEnd()

Public Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Scroll()
Public Event Resize()

Public Enum cnvWorkModeCts
    [cnvScrollMode]
    [cnvPickColorMode]
End Enum

'==================================================================================================

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const RGN_DIFF           As Long = 4
Private Const COLOR_APPWORKSPACE As Long = 12

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'==================================================================================================

Private m_Zoom     As Long
Private m_WorkMode As cnvWorkModeCts
Private m_FitMode  As Boolean
Private m_Enabled  As Boolean

Private m_hPos As Long
Private m_hMax As Long
Private m_vPos As Long
Private m_vMax As Long

Private m_Down As Boolean
Private m_cPt  As POINTAPI

Private m_lsthPos As Single
Private m_lstvPos As Single
Private m_lsthMax As Single
Private m_lstvMax As Single

'==================================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize DIB
    Set DIB = New cDIB32
    
    '-- Default values
    m_Zoom = 1
    m_WorkMode = [cnvScrollMode]
End Sub

Private Sub DIB_Progress(ByVal p As Long)
    RaiseEvent DIBProgress(p)
End Sub

Private Sub DIB_ProgressEnd()
    RaiseEvent DIBProgressEnd
End Sub

'==================================================================================================

Private Sub iCanvas_Paint()
  
  Dim xOff As Long, yOff As Long
  Dim wDst As Long, hDst As Long
  Dim xSrc As Long, ySrc As Long
  Dim wSrc As Long, hSrc As Long
    
    If (DIB.hDIB <> 0) Then
        
        '-- Get Left and Width of source image rectangle:
        If (m_hMax And m_FitMode = 0) Then
            xOff = -m_hPos Mod m_Zoom
            wDst = (iCanvas.Width \ m_Zoom) * m_Zoom + 2 * m_Zoom
            xSrc = m_hPos \ m_Zoom
            wSrc = iCanvas.Width \ m_Zoom + 2
          Else
            xOff = 0
            wDst = iCanvas.Width
            xSrc = 0
            wSrc = DIB.Width
        End If
        '-- Get Top and Height of source image rectangle:
        If (m_vMax And m_FitMode = 0) Then
            yOff = -m_vPos Mod m_Zoom
            hDst = (iCanvas.Height \ m_Zoom) * m_Zoom + 2 * m_Zoom
            ySrc = m_vPos \ m_Zoom
            hSrc = iCanvas.Height \ m_Zoom + 2
          Else
            yOff = 0
            hDst = iCanvas.Height
            ySrc = 0
            hSrc = DIB.Height
        End If
        '-- Paint visible source rectangle:
        DIB.Stretch iCanvas.hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc
    End If
End Sub

Public Sub Repaint()
    iCanvas_Paint
End Sub

Public Sub Resize()
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    
  Dim rW As Long, rH As Long
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
  
    With DIB
        
        If (.hDIB <> 0) Then
        
            If (m_FitMode = 0) Then
            
                '-- Get new Width
                If (.Width * m_Zoom > ScaleWidth) Then
                    m_hMax = .Width * m_Zoom - ScaleWidth
                    rW = ScaleWidth
                  Else
                    m_hMax = 0
                    rW = .Width * m_Zoom
                End If
                '-- Get new Height
                If (.Height * m_Zoom > ScaleHeight) Then
                    m_vMax = .Height * m_Zoom - ScaleHeight
                    rH = ScaleHeight
                  Else
                    m_vMax = 0
                    rH = .Height * m_Zoom
                End If
                
              Else
                DIB.GetBestFitInfo ScaleWidth, ScaleHeight, bfx, bfy, bfW, bfH
            End If
            
            '-- Resize
            If (m_FitMode = 0) Then
                MoveWindow iCanvas.hwnd, (ScaleWidth - rW) \ 2, (ScaleHeight - rH) \ 2, rW, rH, 0
              Else
                MoveWindow iCanvas.hwnd, bfx, bfy, bfW, bfH, 0
            End If
                                
            '-- Memory position:
            '   Horizontal position
                If (m_lsthMax) Then
                    m_hPos = m_lsthPos * m_hMax / m_lsthMax
                  Else
                    m_hPos = m_hMax / 2
                End If
            '   Vertical position
                If (m_lstvMax) Then
                    m_vPos = m_lstvPos * m_vMax / m_lstvMax
                  Else
                    m_vPos = m_vMax / 2
                End If
            '   Save values
                m_lsthPos = m_hPos: m_lstvPos = m_vPos
                m_lsthMax = m_hMax: m_lstvMax = m_vMax
            
            '-- Refresh
            pvCls
            Repaint
            
            '-- Raise Resize event
            RaiseEvent Resize
          
          Else
            '-- Cls (whole)
            iCanvas.Move -1, -1, 0, 0
            pvCls
        End If
    End With
    
    '-- Reset pointer
    iCanvas.MouseIcon = Nothing
    '-- Change it
    If (m_WorkMode = [cnvScrollMode]) Then
        If ((m_hMax Or m_vMax) And Not m_FitMode) Then
            iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDOVER", vbResCursor)
        End If
      Else
        iCanvas.MouseIcon = LoadResPicture("CURSOR_PICKCOLOR", vbResCursor)
    End If
End Sub

Private Sub pvCls()

  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim hBrush As Long
    
    '-- Create brush (background)
    hBrush = GetSysColorBrush(COLOR_APPWORKSPACE)
    
    '-- Create Cls region (Control Rect. - iCanvas Rect.)
    With iCanvas
        hRgn_1 = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
        hRgn_2 = CreateRectRgn(.Left, .Top, .Left + .Width - 1, .Top + .Height - 1)
    End With
    CombineRgn hRgn_1, hRgn_1, hRgn_2, RGN_DIFF
    
    '-- Fill it
    FillRgn hDC, hRgn_1, hBrush
    
    '-- Clear
    DeleteObject hBrush
    DeleteObject hRgn_1
    DeleteObject hRgn_2
End Sub

'==================================================================================================

Public Property Let Zoom(ByVal Factor As Long)
    m_Zoom = IIf(Factor < 1, 1, Factor)
End Property

Public Property Get Zoom() As Long
Attribute Zoom.VB_MemberFlags = "400"
    Zoom = m_Zoom
End Property

Public Property Let WorkMode(ByVal Mode As cnvWorkModeCts)

    '-- Reset pointer
    iCanvas.MouseIcon = Nothing
    
    '-- Change it
    If (Mode = [cnvScrollMode]) Then
        If ((m_hMax Or m_vMax) And Not m_FitMode) Then
            iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDOVER", vbResCursor)
        End If
      Else
        iCanvas.MouseIcon = LoadResPicture("CURSOR_PICKCOLOR", vbResCursor)
    End If
    m_WorkMode = Mode
End Property

Public Property Get WorkMode() As cnvWorkModeCts
Attribute WorkMode.VB_MemberFlags = "400"
    WorkMode = m_WorkMode
End Property

Public Property Let FitMode(ByVal Enable As Boolean)
    m_FitMode = Enable
End Property

Public Property Get FitMode() As Boolean
Attribute FitMode.VB_MemberFlags = "400"
    FitMode = m_FitMode
End Property

Public Property Let Enabled(ByVal Enable As Boolean)
    UserControl.Enabled = Enable
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

'==================================================================================================

Private Sub iCanvas_DblClick()
    iCanvas_MouseDown 1, 0, CSng(m_cPt.x), CSng(m_cPt.y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    '-- iCanvas offset
    x = x - iCanvas.Left
    y = y - iCanvas.Top
    
    '-- Set mouse capture to iCanvas
    SetCapture iCanvas.hwnd
    iCanvas_MouseDown Button, Shift, x, y
End Sub

Private Sub iCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Change pointer
    If ((m_hMax Or m_vMax) And m_WorkMode = [cnvScrollMode] And Not m_FitMode) Then
        iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDCATCH", vbResCursor)
    End If
    
    m_Down = (Button = 1)
    m_cPt.x = x
    m_cPt.y = y
    
    RaiseEvent MouseDown(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub iCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If (m_Down) Then
    
        Select Case m_WorkMode
        
          Case [cnvScrollMode]
            '-- Get displacements
            m_hPos = m_hPos + (m_cPt.x - x)
            m_vPos = m_vPos + (m_cPt.y - y)
            '-- Check margins
            If (m_hPos < 0) Then m_hPos = 0
            If (m_vPos < 0) Then m_vPos = 0
            If (m_hPos > m_hMax) Then m_hPos = m_hMax
            If (m_vPos > m_vMax) Then m_vPos = m_vMax
            '-- Save current position
            m_cPt.x = x
            m_cPt.y = y
            
            If (m_lsthPos <> m_hPos Or m_lstvPos <> m_vPos) Then
                '-- Refresh
                Repaint
                '-- Raise Scroll event
                RaiseEvent Scroll
            End If
            m_lsthPos = m_hPos
            m_lstvPos = m_vPos
        
        End Select
    End If
    
    RaiseEvent MouseMove(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub iCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Change pointer
    If ((m_hMax Or m_vMax) And WorkMode = [cnvScrollMode] And Not m_FitMode) Then
        iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDOVER", vbResCursor)
    End If
    
    m_Down = False
    
    RaiseEvent MouseUp(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub iCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Not m_Down) Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'==================================================================================================

Private Function pvDIBx(ByVal x As Long) As Long
    pvDIBx = Int((IIf(m_FitMode, 0, m_hPos) + x) / m_Zoom)
End Function

Private Function pvDIBy(ByVal y As Long) As Long
    pvDIBy = Int((IIf(m_FitMode, 0, m_vPos) + y) / m_Zoom)
End Function

