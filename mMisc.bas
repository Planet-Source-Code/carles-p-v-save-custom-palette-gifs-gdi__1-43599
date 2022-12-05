Attribute VB_Name = "mMisc"
Option Explicit

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private m_hBrush As Long



Public Sub InitializePatternBrush()

  Dim hBitmap        As Long
  Dim tBytes(1 To 8) As Integer
    
    '-- Brush pattern (8x8)
    tBytes(1) = &HAA
    tBytes(2) = &H55
    tBytes(3) = &HAA
    tBytes(4) = &H55
    tBytes(5) = &HAA
    tBytes(6) = &H55
    tBytes(7) = &HAA
    tBytes(8) = &H55
    
    '-- Create brush
    hBitmap = CreateBitmap(8, 8, 1, 1, tBytes(1))
    m_hBrush = CreatePatternBrush(hBitmap)
    DeleteObject hBitmap
End Sub

Public Sub DestroyPatternBrush()
    DeleteObject m_hBrush
End Sub

Public Function DrawRectangle(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)

  Dim rRect  As RECT2
  Dim hBrush As Long
  
    If (lColor > -1) Then
        hBrush = CreateSolidBrush(lColor)
        SetRect rRect, x1, y1, x2, y2
        FillRect hDC, rRect, hBrush
        DeleteObject hBrush
      Else
        SetRect rRect, x1, y1, x2, y2
        FillRect hDC, rRect, m_hBrush
    End If
End Function

Public Sub RemoveButtonBorderEnhancement(Button As CommandButton)
    SendMessage Button.hwnd, &HF4&, &H0&, 0&
End Sub
