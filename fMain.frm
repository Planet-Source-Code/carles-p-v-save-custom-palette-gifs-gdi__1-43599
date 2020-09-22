VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Custom GIF palette (GDI+)"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCountUniqueColors 
      Appearance      =   0  'Flat
      Caption         =   "&Unique colors:"
      Height          =   210
      Left            =   195
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6075
      Width           =   1335
   End
   Begin VB.CheckBox chkPickColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Pick &color"
      Enabled         =   0   'False
      Height          =   210
      Left            =   4380
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5610
      Width           =   975
   End
   Begin VB.Frame fraDither 
      Caption         =   "Dithering"
      Enabled         =   0   'False
      Height          =   5385
      Left            =   5610
      TabIndex        =   18
      Top             =   2490
      Width           =   2955
      Begin VB.VScrollBar sbHalftoneLevels 
         Height          =   285
         Left            =   2385
         Max             =   2
         Min             =   6
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1425
         Value           =   5
         Width           =   270
      End
      Begin VB.TextBox txtHalftoneLevels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "125"
         Top             =   1425
         Width           =   465
      End
      Begin VB.VScrollBar sbNoise 
         Height          =   285
         Left            =   2070
         Max             =   0
         Min             =   25
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   645
         Width           =   270
      End
      Begin VB.TextBox txtNoise 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1605
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   645
         Width           =   465
      End
      Begin VB.CheckBox chkNoise 
         Appearance      =   0  'Flat
         Caption         =   "Pre-apply noise"
         Height          =   300
         Left            =   150
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   645
         Width           =   1440
      End
      Begin VB.CommandButton cmdResetWeights 
         Height          =   225
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3495
         Width           =   240
      End
      Begin VB.OptionButton optPalette 
         Appearance      =   0  'Flat
         Caption         =   "Halftone:"
         Height          =   240
         Index           =   3
         Left            =   1665
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1095
      End
      Begin VB.OptionButton optPalette 
         Appearance      =   0  'Flat
         Caption         =   "Websave"
         Height          =   240
         Index           =   4
         Left            =   1665
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin VB.VScrollBar sbOptimalEntriesFast 
         Height          =   285
         Left            =   2055
         Max             =   8
         Min             =   256
         SmallChange     =   10
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2460
         Value           =   256
         Width           =   270
      End
      Begin VB.CommandButton cmdUseDefaultWeights 
         Height          =   225
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3210
         Width           =   240
      End
      Begin VB.CheckBox chkUseTransparent 
         Appearance      =   0  'Flat
         Caption         =   "Use transparent entry (Pick)"
         Height          =   270
         Left            =   180
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4725
         Width           =   2370
      End
      Begin VB.CommandButton cmdSaveGIF 
         Caption         =   "&Save Test.gif"
         Enabled         =   0   'False
         Height          =   450
         Left            =   1515
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   4170
         Width           =   1245
      End
      Begin VB.CommandButton cmdDither 
         Caption         =   "&Dither"
         Enabled         =   0   'False
         Height          =   450
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   4170
         Width           =   1245
      End
      Begin VB.OptionButton optPalette 
         Appearance      =   0  'Flat
         Caption         =   "Black and White"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1650
      End
      Begin VB.OptionButton optPalette 
         Appearance      =   0  'Flat
         Caption         =   "16 greys"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1485
         Width           =   1650
      End
      Begin VB.OptionButton optPalette 
         Appearance      =   0  'Flat
         Caption         =   "256 greys"
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1650
      End
      Begin VB.OptionButton optPalette 
         Appearance      =   0  'Flat
         Caption         =   "Optimal (Octree):"
         Height          =   240
         Index           =   5
         Left            =   150
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2115
         Width           =   1650
      End
      Begin VB.TextBox txtOptimalEntries 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "256"
         Top             =   2460
         Width           =   465
      End
      Begin VB.CheckBox chkDiffuseError 
         Appearance      =   0  'Flat
         Caption         =   "Diffuse error (FS)"
         Height          =   300
         Left            =   150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   315
         Width           =   1590
      End
      Begin VB.TextBox txtRGBWeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   2820
         Width           =   465
      End
      Begin VB.TextBox txtRGBWeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   2820
         Width           =   465
      End
      Begin VB.TextBox txtRGBWeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   2820
         Width           =   465
      End
      Begin VB.VScrollBar sbOptimalEntries 
         Height          =   285
         Left            =   1785
         Max             =   8
         Min             =   256
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2460
         Value           =   256
         Width           =   270
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   180
         X2              =   2745
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   180
         X2              =   2745
         Y1              =   3945
         Y2              =   3945
      End
      Begin VB.Label lblResetWeights 
         Caption         =   "Reset weights"
         Height          =   210
         Left            =   660
         TabIndex        =   42
         Top             =   3510
         Width           =   2100
      End
      Begin VB.Label lblDefaultWeights 
         Caption         =   "Use correction weights (def.)"
         Height          =   210
         Left            =   660
         TabIndex        =   40
         Top             =   3225
         Width           =   2100
      End
      Begin VB.Label lblFileSize 
         Height          =   210
         Left            =   900
         TabIndex        =   47
         Top             =   5025
         Width           =   1455
      End
      Begin VB.Label lblFileSizeT 
         Caption         =   "File size:"
         Height          =   240
         Left            =   195
         TabIndex        =   46
         Top             =   5025
         Width           =   645
      End
      Begin VB.Label lblOptimalEntries 
         Caption         =   "Max. entries"
         Height          =   255
         Left            =   330
         TabIndex        =   31
         Top             =   2490
         Width           =   990
      End
      Begin VB.Label lblRGBWeights 
         Caption         =   "RGB weights"
         Height          =   255
         Left            =   330
         TabIndex        =   35
         Top             =   2850
         Width           =   1020
      End
   End
   Begin VB.TextBox lblZoom 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   5235
      Width           =   675
   End
   Begin VB.VScrollBar sbZoom 
      Height          =   285
      Left            =   5085
      Max             =   1
      Min             =   10
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5235
      Value           =   1
      Width           =   270
   End
   Begin VB.CommandButton cmdPasteFromClipboard 
      Caption         =   "&Paste from clipboard"
      Height          =   450
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7410
      Width           =   2025
   End
   Begin VB.PictureBox iPalette 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   1995
      Left            =   5625
      MousePointer    =   99  'Custom
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   375
      Width           =   1995
      Begin VB.Shape shpSelect 
         BorderColor     =   &H00FFFFFF&
         Height          =   135
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load &GrapeBunch.bmp"
      Height          =   450
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6855
      Width           =   2025
   End
   Begin GIF.ucCanvas ucPreview 
      Height          =   4770
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8414
   End
   Begin VB.Label lblBitmapSizeT 
      Caption         =   "Bitmap file size:"
      Height          =   240
      Left            =   195
      TabIndex        =   6
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Label lblBitmapSize 
      Caption         =   " "
      Height          =   240
      Left            =   1560
      TabIndex        =   7
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblDimensionsT 
      Caption         =   "Dimensions:"
      Height          =   240
      Left            =   195
      TabIndex        =   2
      Top             =   5250
      Width           =   1095
   End
   Begin VB.Label lblDimensions 
      Height          =   240
      Left            =   1560
      TabIndex        =   3
      Top             =   5250
      Width           =   825
   End
   Begin VB.Label lblZoomT 
      Alignment       =   1  'Right Justify
      Caption         =   "Zoom"
      Height          =   210
      Left            =   3825
      TabIndex        =   10
      Top             =   5265
      Width           =   525
   End
   Begin VB.Label lblbpp 
      Caption         =   " "
      Height          =   240
      Left            =   1560
      TabIndex        =   5
      Top             =   5505
      Width           =   2055
   End
   Begin VB.Label lblbppT 
      Caption         =   "Color depth:"
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   5505
      Width           =   990
   End
   Begin VB.Label lblCurrentPalette 
      Caption         =   "Current palette"
      Height          =   270
      Left            =   5640
      TabIndex        =   16
      Top             =   135
      Width           =   1545
   End
   Begin VB.Label lblUniqueColors 
      Height          =   240
      Left            =   1560
      TabIndex        =   9
      Top             =   6075
      Width           =   825
   End
   Begin VB.Label lblImagePreview 
      Caption         =   "Image preview"
      Height          =   270
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// --------------------------------------------
'// Saving GIFs with custom palette through GDI+
'// --------------------------------------------
'//
'// Thanks to vbAccelerator, specialy for:
'//
'//     cPalette.cls: Steve McMahon
'//     Octree Colour Quantisation Code (CreateOptimal): Brian Schimpf
'//
'// Thanks to Avery for helping me about GDI+ GIF issues and for his
'// great GDI+ submission:
'// http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'//
'//
'// Notes (optimal palette):
'//
'//   - We'll get good results from 32+ entries: more colors, bigger file size.
'//
'//   - The extracion of the optimal palette is an expensive algorithm.
'//     Two accelarations have been done here:
'//     1. Size reduction of source DIB.
'//     2. Reduction of original RGB space ("True color") to a 4096-colors one:
'//        See cDither8bpp.DitherToColorPalette sub.
'//        Grey dithering is a more simple case. Palette index corresponds to source L value.
'//
'//   - RGB weights: Octree method works well for dithering to 8 bpp (256 colors). Extraction
'//     of optimal palettes with a low number of entries will tend to 'decontrast' resulting
'//     palette. These weights can significantly correct it.
'//     Testing several weights, next ones, seem to work quite well:
'//
'//         wR = 0.360, wG = 0.436, wB = 0.341
'//
'//     In fact, these numbers come from their respective L weight according:
'//
'//         wChannel = f(Lchannel)
'//         wR = 1 / (3 - 0.222) = 0.360
'//         wG = 1 / (3 - 0.707) = 0.436
'//         wB = 1 / (3 - 0.071) = 0.341
'//
'// More info:
'//
'//   - http://www.mactech.com/articles/develop/issue_10/Good_Othmer_final.html
'//   - http://www.eg.org/EG/CGF/volume17/issue3/269.pdf
'//

'// Log:
'//
'// - 2003.03.01:
'//   Fixed: Websafe palette generation (offset at last static entries).
'//   Added: Halftone levels. Now: 8, 27, 64, 125, 216 levels.
'//   Imprv: Final gif size (passing real palette entries instead of ct. 256).
'//
'// - 2003.03.03:
'//   Fixed: Halftone-64 (incorrect level step: &H50, should be &H55).
'//
'// - 2003.03.13:
'//   IMPVD: Removed DIB buffer copy from Dithering functions.



Option Explicit

Private DIBPalette As New cPalette8bpp '-- 8 bpp palette
Private DIBDither  As New cDither8bpp  '-- 8 bpp dithering

Private m_DIB As New cDIB32            '-- Temp. copy of our DIB (32 bpp)
Private m_bpp As Byte                  '-- Current color depth
Private m_PaletteTypeIndex As Integer  '-- Current palette type

Private m_GDIpToken As Long            '-- Needed to open/close GDI+






Private Sub Form_Load()

  Dim GpInput As GdiplusStartupInput
  
    '-- In IDE?
    If (App.LogMode <> 1) Then
        MsgBox "Run compiled. Strongly recomended", vbInformation
    End If
    
    '-- Load the GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        MsgBox "Error loading GDI+!", vbCritical
        Unload Me
        Exit Sub
    End If
    
    '-- Initialize palette (blank 256-colors palette)
    pv_Initialize
    
    '-- Initialize pattern brush (unused palette entries)
    InitializePatternBrush
    
    '-- Remove buttons border enhancement
    RemoveButtonBorderEnhancement cmdLoad
    RemoveButtonBorderEnhancement cmdPasteFromClipboard
    RemoveButtonBorderEnhancement cmdDither
    RemoveButtonBorderEnhancement cmdSaveGIF
    RemoveButtonBorderEnhancement cmdUseDefaultWeights
    RemoveButtonBorderEnhancement cmdResetWeights
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '-- Unload the GDI+ Dll
    GdiplusShutdown m_GDIpToken
    
    '-- Destroy pattern brush
    DestroyPatternBrush
    
    '-- Free objects
    ucPreview.DIB.Destroy
    Set m_DIB = Nothing
    Set DIBPalette = Nothing
    Set DIBDither = Nothing
    Set fMain = Nothing
End Sub

Private Sub Form_Paint()
    Line (ucPreview.Left, 429)-(ucPreview.Left + ucPreview.Width, 429), vb3DShadow
    Line (ucPreview.Left, 430)-(ucPreview.Left + ucPreview.Width, 430), vb3DHighlight
End Sub


'==================================================================================================
' Load test bitmap (GrapeBunch.bmp) / Paste from clipboard / Zoom / Pick color / Unique colors
'==================================================================================================

Private Sub cmdLoad_Click()
    '-- Load test bitmap
    pv_SetDIB LoadPicture(App.Path & "\GrapeBunch.bmp")
End Sub

Private Sub cmdPasteFromClipboard_Click()
    '-- Get from clipboard
    If (Clipboard.GetFormat(vbCFBitmap)) Then
        pv_SetDIB Clipboard.GetData(vbCFBitmap)
      Else
        MsgBox "Empty clipboard", vbInformation
    End If
End Sub



Private Sub sbZoom_Change()
    '-- Change DIB preview zoom
    ucPreview.Zoom = sbZoom
    ucPreview.Resize
    lblZoom = Format(ucPreview.Zoom, "0%")
End Sub

Private Sub chkPickColor_Click()
    '-- Enable/disable <Pick color> mode
    If (chkPickColor = 1) Then
        ucPreview.WorkMode = [cnvPickColorMode]
      Else
        ucPreview.WorkMode = [cnvScrollMode]
    End If
End Sub

Private Sub chkCountUniqueColors_Click()
    '-- Enable/disable unique colors counting
    If (chkCountUniqueColors And ucPreview.DIB.hDIB <> 0) Then
        Screen.MousePointer = vbArrowHourglass
        lblUniqueColors = Format(DIBDither.CountColors(ucPreview.DIB), "#,#")
        Screen.MousePointer = vbDefault
      Else
        lblUniqueColors = ""
    End If
End Sub


'==================================================================================================
' Select palette
'==================================================================================================

Private Sub optPalette_Click(Index As Integer)

  Dim sDIB As New cDIB32
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
   
    '-- Save palette type (dither)
    m_PaletteTypeIndex = Index
   
    '-- Build palette
    Select Case Index
      Case 0
        DIBPalette.CreateGreyscale [002pgLevels]
      Case 1
        DIBPalette.CreateGreyscale [016pgLevels]
      Case 2
        DIBPalette.CreateGreyscale [256pgLevels]
      Case 3
        Select Case sbHalftoneLevels.Value
          Case 2
            DIBPalette.CreateHalftone [008phLevels]
          Case 3
            DIBPalette.CreateHalftone [027phLevels]
          Case 4
            DIBPalette.CreateHalftone [064phLevels]
          Case 5
            DIBPalette.CreateHalftone [125phLevels]
          Case 6
            DIBPalette.CreateHalftone [216phLevels]
        End Select
      Case 4
        DIBPalette.CreateWebsafe
      Case 5
        '-- Speed up: Reduce source DIB from which we are going to
        '   extract optimal palette (fit to 100x100... enough).
        ucPreview.DIB.GetBestFitInfo 100, 100, bfx, bfy, bfW, bfH
        sDIB.Create bfW, bfH
        sDIB.LoadDIBBlt m_DIB
        '-- Get optimal palette
        Screen.MousePointer = vbArrowHourglass
        DIBPalette.CreateOptimal sDIB, Val(txtOptimalEntries), 8, Val(txtRGBWeight(0)) / 1000, Val(txtRGBWeight(1)) / 1000, Val(txtRGBWeight(2)) / 1000
        Screen.MousePointer = vbDefault
    End Select
    
    '-- Draw palette
    pv_DrawPalette
    
    '-- Enable dither/Disable Save/Palette cursor
    cmdDither.Enabled = -1
    cmdSaveGIF.Enabled = 0
    shpSelect.Visible = 0
    chkPickColor.Enabled = 0: chkPickColor = 0
End Sub


'-- Halftone palette levels

Private Sub sbHalftoneLevels_Change()
    '-- Levels (channel steps ^ 3)
    txtHalftoneLevels = sbHalftoneLevels ^ 3
    '-- Force Halftone selection
    If (optPalette(3) = 0) Then
        optPalette(3) = 1
      Else
        optPalette_Click 3
    End If
End Sub


'-- Optimal palette settings

Private Sub chkDiffuseError_Click()
    '-- Disable Save
    cmdSaveGIF.Enabled = 0
End Sub

Private Sub chkNoise_Click()
    '-- Disable Save
    cmdSaveGIF.Enabled = 0
End Sub

Private Sub sbNoise_Change()
    txtNoise = sbNoise
    '-- Disable Save
    cmdSaveGIF.Enabled = 0
End Sub

Private Sub txtOptimalEntries_Change()
    pv_Initialize
End Sub

Private Sub sbOptimalEntries_Change()
    txtOptimalEntries = sbOptimalEntries
    sbOptimalEntriesFast = sbOptimalEntries
End Sub
Private Sub sbOptimalEntriesFast_Change()
    sbOptimalEntries = sbOptimalEntriesFast
End Sub

Private Sub txtRGBWeight_Change(Index As Integer)

    '-- Check range [0-1000]
    With txtRGBWeight(Index)
        If (Val(.Text) > 1000) Then
            .Text = 1000
            .SelStart = 0
            .SelLength = 4
        End If
    End With
    
    pv_Initialize
End Sub

Private Sub txtRGBWeight_KeyPress(Index As Integer, KeyAscii As Integer)
    If ((KeyAscii < 47 Or KeyAscii > 58) And KeyAscii <> 8) Then KeyAscii = 0
End Sub

Private Sub cmdUseDefaultWeights_Click()
    txtRGBWeight(0) = 360
    txtRGBWeight(1) = 436
    txtRGBWeight(2) = 341
End Sub

Private Sub cmdResetWeights_Click()
    txtRGBWeight(0) = 1000
    txtRGBWeight(1) = 1000
    txtRGBWeight(2) = 1000
End Sub


'==================================================================================================
' Dither to current palette
'==================================================================================================

Private Sub cmdDither_Click()
    
  Dim DIBFilter As New cDIB32Filter
    
    Screen.MousePointer = vbArrowHourglass
    
    '-- Restore to original
    ucPreview.DIB.LoadBlt m_DIB.hDIBDC
    
    '-- Pre-apply noise ?
    If (chkNoise And Val(txtNoise) > 0) Then
        DIBFilter.Noise ucPreview.DIB, Val(txtNoise)
    End If
    
    '-- Dither
    Select Case m_PaletteTypeIndex
      Case 0
        DIBDither.DitherToGreyPalette ucPreview.DIB, DIBPalette, [002dgLevels], CBool(chkDiffuseError)
      Case 1
        DIBDither.DitherToGreyPalette ucPreview.DIB, DIBPalette, [016dgLevels], CBool(chkDiffuseError)
      Case 2
        DIBDither.DitherToGreyPalette ucPreview.DIB, DIBPalette, [256dgLevels], CBool(chkDiffuseError)
      Case 3, 4, 5
        DIBDither.DitherToColorPalette ucPreview.DIB, DIBPalette, CBool(chkDiffuseError)
    End Select
    
    '-- Refresh view
    ucPreview.Repaint
    
    '-- Change color depth
    Select Case DIBPalette.MaxCount
      Case Is <= 2
        m_bpp = 1
      Case Is <= 16
        m_bpp = 4
      Case Is <= 256
        m_bpp = 8
    End Select
    lblbpp = m_bpp & " bpp (virtual color depth)"
    
    '-- Show bitmap file size (header + palette + bitmap bits)
    lblBitmapSize = Format(54 + 4 * 2 ^ m_bpp + (((ucPreview.DIB.Width * m_bpp + 31) \ 32) * 4) * ucPreview.DIB.Height, "#,# bytes") & " (" & m_bpp & " bpp)"
    
    '-- Count unique colors and "change" color depth
    If (chkCountUniqueColors) Then
        lblUniqueColors = Format(DIBDither.CountColors(ucPreview.DIB), "#,#")
    End If
    
    '-- Enable Save/Pick color mode/Palette cursor
    cmdSaveGIF.Enabled = -1
    chkPickColor.Enabled = -1
    shpSelect.Visible = -1: shpSelect.Move 0, 0
    
    Screen.MousePointer = vbDefault
End Sub


'==================================================================================================
' Save Test.gif
'==================================================================================================

Private Sub cmdSaveGIF_Click()

  Dim iTrnspEntry As Integer
    
    '-- Get current selected palette entry
    If (chkUseTransparent = 1) Then
        iTrnspEntry = (shpSelect.Left \ 8) + 16 * (shpSelect.Top \ 8)
      Else
        iTrnspEntry = -1
    End If

    If (m_bpp <= 8) Then
    
        '-- Save "Test.gif"
        SaveDIB256ToGIF ucPreview.DIB, _
                        DIBPalette, _
                        DIBDither, _
                        iTrnspEntry, _
                        App.Path & "\Test.gif"
                        
        '-- Show file size
        lblFileSize = Format(FileLen(App.Path & "\Test.gif"), "#,# bytes")
    End If
End Sub


'==================================================================================================
' Palette entry selection (transparent color)
'==================================================================================================

Private Sub ucPreview_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    ucPreview_MouseMove vbLeftButton, 0, x, y
End Sub

Private Sub ucPreview_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    
  Dim PalIndex As Byte
    
    '-- Check if dithered...
    If (Button = vbLeftButton And (m_bpp <= 8 And chkPickColor)) Then
        
        '-- Check out of bounds...
        If (x >= 0 And y >= 0 And x < ucPreview.DIB.Width And y < ucPreview.DIB.Height) Then
    
            '-- Get palette index
            PalIndex = DIBDither.PalID(x, y)
            
            '-- Search color in palette (update cursor)
            shpSelect.Move (PalIndex Mod 16) * 8, (PalIndex \ 16) * 8
        End If
    End If
End Sub







'==================================================================================================
' Private
'==================================================================================================

Private Sub pv_SetDIB(Picture As StdPicture)
    
    '-- Load and get original color depth. Initialize
    m_bpp = ucPreview.DIB.CreateFromStdPicture(Picture)
    ucPreview.Resize
    pv_Initialize
    
    '-- Create a temp. copy
    m_DIB.Create ucPreview.DIB.Width, ucPreview.DIB.Height
    m_DIB.LoadBlt ucPreview.DIB.hDIBDC
    
    '-- Show original bitmap dimensions
    lblDimensions = ucPreview.DIB.Width & " x " & ucPreview.DIB.Height
    '-- Show original bitmap color depth (32 bpp)
    lblbpp = m_bpp & " bpp (memory DIB)"
    '-- Show bitmap file size (24 bpp) [54 = BITMAPFILEHEADER + BITMAPINFOHEADER]
    lblBitmapSize = Format(54 + (((ucPreview.DIB.Width * 24 + 31) \ 32) * 4) * ucPreview.DIB.Height, "#,# bytes (24 bpp)")

    '-- Count original unique colors ?
    If (chkCountUniqueColors) Then
        Screen.MousePointer = vbArrowHourglass
        lblUniqueColors = Format(DIBDither.CountColors(ucPreview.DIB), "#,#")
        Screen.MousePointer = vbDefault
    End If
    
    '-- Success...
    fraDither.Enabled = (m_bpp <> 0)
End Sub

Private Sub pv_DrawPalette()
  
  Dim i As Long, j As Long
  Dim lIndex As Long
  Dim lColor As Long
    
    '-- Show the 256 entries
    For i = 0 To 120 Step 8
        For j = 0 To 120 Step 8
            With DIBPalette
                If (lIndex < DIBPalette.MaxCount) Then
                    lColor = RGB(.rgbR(lIndex), .rgbG(lIndex), .rgbB(lIndex))
                  Else
                    lColor = -1
                End If
                DrawRectangle iPalette.hDC, j + 1, i + 1, j + 8, i + 8, lColor
                lIndex = lIndex + 1
            End With
        Next j
    Next i
    iPalette.Refresh
End Sub

Private Sub pv_Initialize()

    '-- Initiliaze palette
    DIBPalette.InitializePalette
    pv_DrawPalette
    '-- Unselect
    optPalette(m_PaletteTypeIndex).Value = 0
    
    '-- Disable Dither/Save/Pick color mode/Palette cursor
    cmdDither.Enabled = 0
    cmdSaveGIF.Enabled = 0
    chkPickColor.Enabled = 0: chkPickColor = 0
    shpSelect.Visible = 0
End Sub
