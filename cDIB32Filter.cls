VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB32Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' cDIB32Filter.cls (Removed all other filters)
' Dependencies: cDIB32.cls

Option Explicit

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

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

Private t As Long
Private x As Long, xIn As Long
Private y As Long, yIn As Long
Private W As Long
Private H As Long

'==================================================================================================

Public Event Progress(ByVal p As Long)
Public Event ProgressEnd()

'==================================================================================================

Public Sub Noise(DIB As cDIB32, Optional ByVal Level As Long = 5, Optional ByVal xyStep As Long = 1)
    
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim Lev As Long, DblLev As Long
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
    
    If (DIB.hDIB <> 0) Then
    
        pvBuildSA tSA, DIB
        CopyMemory ByVal VarPtrArray(Bits()), VarPtr(tSA), 4
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        DblLev = 2 * Level
        
        For y = 0 To H Step xyStep
            For x = 0 To W Step xyStep
                Lev = Rnd * DblLev - Level
                B = CLng(Bits(x, y).B) + Lev
                G = CLng(Bits(x, y).G) + Lev
                R = CLng(Bits(x, y).R) + Lev
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With Bits(x, y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next x
'            RaiseEvent Progress(y)
        Next y
        CopyMemory ByVal VarPtrArray(Bits), 0&, 4
        RaiseEvent ProgressEnd
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
