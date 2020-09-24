Type FontLog
     FontHeight As Integer
     FontWidth As Integer
     FontRotate As Integer
     FontOrientation As Integer
     FontWeight As Integer
     FontItalic As String * 1
     FontUnderLine As String * 1
     FontStrikeThru As String * 1
     FontCharSet As String * 1
     FontOutPrecision As String * 1
     FontClipPrecision As String * 1
     FontQuality As String * 1
     FontPitchFamily As String * 1
     FontFaceName As String * 20
End Type
''''''''
Type TEXTMETRIC
     tmHeight As Integer
     tmAscent As Integer
     tmDescent As Integer
     tmInternalLeading As Integer
     tmExternalLeading As Integer
     tmAveCharWidth As Integer
     tmMaxCharWidth As Integer
     tmWeight As Integer
     tmItalic As String * 1
     tmUnderlined As String * 1
     tmStruckOut As String * 1
     tmFirstChar As String * 1
     tmLastChar As String * 1
     tmDefaultChar As String * 1
     tmBreakChar As String * 1
     tmPitchAndFamily As String * 1
     tmCharSet As String * 1
     tmOverhang As Integer
     tmDigitizedAspectX As Integer
     tmDigitizedAspectY As Integer
End Type
''''''''
Type POINTAPI
     x As Integer
     y As Integer
End Type
''''''''
Type GLYPHMETRICS
     gmBlackBoxX As Integer
     gmBlackBoxY As Integer
     gmptGlyphOrigin As POINTAPI
     gmCellIncX As Integer
     gmCellIncY As Integer
End Type
''''''''
Type Fixed
   fract As Integer
   value As Integer
End Type
''''''''
Type Mat2
   eM11 As Fixed
   eM12 As Fixed
   eM21 As Fixed
   eM22 As Fixed
End Type
''''''''
Type POINTFX
   x As Fixed
   y As Fixed
End Type
''''''''
Type TTPOLYGONHEADER
   cb As Long
   dw As Long
   pfxStart As POINTFX
End Type
''''''''
Type TTPOLYCURVE
   wType As Integer
   cpfx As Integer
   pfx(0 To 2) As POINTFX
End Type
''''''''
Declare Function CreateFontIndirect% Lib "GDI" (lpFontLog As FontLog)
Declare Function DeleteObject% Lib "GDI" (ByVal hObject%)
Declare Function GetTextMetrics% Lib "GDI" (ByVal hDC%, lpMetrics As TEXTMETRIC)
Declare Function GetTextFace% Lib "GDI" (ByVal hDC%, ByVal nCount%, ByVal lpFacename$)
Declare Function GetGlyphOutLine& Lib "GDI" (ByVal hDC%, ByVal uChar%, ByVal fuFormat%, lpgm As GLYPHMETRICS, ByVal cbBuffer&, lppt As POINTAPI, lpmat2 As Mat2)
Declare Function GlobalUnlock% Lib "Kernel" (ByVal hMem%)
Declare Function GlobalFree% Lib "Kernel" (ByVal hMem%)
Declare Function PolyPolygon% Lib "GDI" (ByVal hDC%, lpPoints As POINTAPI, lpPolyCounts%, ByVal nCount%)
Declare Function SelectObject% Lib "GDI" (ByVal hDC%, ByVal hObject%)
Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'''''''''''''''''
Global FontSelection%
'''''''''''''''''
Global Const ANSI_CHARSET = 0

Global Const CLIP_CHARACTER_PRECIS = 1
Global Const CLIP_DEFAULT_PRECIS = 0
Global Const CLIP_EMBEDDED = &H80
Global Const CLIP_LH_ANGLES = &H10
Global Const CLIP_STROKE_PRECIS = 2
Global Const CLIP_TT_ALWAYS = &H20

Global Const DEFAULT_QUALITY = 0
Global Const DEFAULT_CHARSET = 1
Global Const DEFAULT_PITCH = 0

Global Const DEVICE_FONTTYPE = 2

Global Const DRAFT_QUALITY = 1

Global Const DT_TOP = &H0
Global Const DT_LEFT = &H0

Global Const DT_BOTTOM = &H8
Global Const DT_CALCRECT = &H400
Global Const DT_CENTER = &H1
Global Const DT_EXPANDTABS = &H40
Global Const DT_EXTERNALLEADING = &H200
Global Const DT_INTERNAL = &H1000
Global Const DT_NOCLIP = &H100
Global Const DT_NOPREFIX = &H800
Global Const DT_RIGHT = &H2
Global Const DT_SINGLELINE = &H20
Global Const DT_TABSTOP = &H80
Global Const DT_VCENTER = &H4
Global Const DT_WORDBREAK = &H10

Global Const FIXED_PITCH = 1

Global Const FF_DECORATIVE = 80
Global Const FF_DONTCARE = 0
Global Const FF_MODERN = 48
Global Const FF_ROMAN = 16
Global Const FF_SCRIPT = 64
Global Const FF_SWISS = 32

Global Const FW_BOLD = 700
Global Const FW_DONTCARE = 0
Global Const FW_EXTRABOLD = 800
Global Const FW_EXTRALIGHT = 200
Global Const FW_HEAVY = 900
Global Const FW_LIGHT = 300
Global Const FW_MEDIUM = 500
Global Const FW_NORMAL = 400
Global Const FW_SEMIBOLD = 600
Global Const FW_THIN = 100

Global Const LF_FULLFACESIZE = 64

Global Const NTM_BOLD = &H20&
Global Const NTM_ITALIC = &H1&
Global Const NTM_REGULAR = &H40&

Global Const OEM_CHARSET = 255

Global Const OUT_CHARACTER_PRECIS = 2
Global Const OUT_DEFAULT_PRECIS = 0
Global Const OUT_DEVICE_PRECIS = 5
Global Const OUT_RASTER_PRECIS = 6
Global Const OUT_STRING_PRECIS = 1
Global Const OUT_STROKE_PRECIS = 3
Global Const OUT_TT_ONLY_PRECIS = 7
Global Const OUT_TT_PRECIS = 4

Global Const PROOF_QUALITY = 2

Global Const RASTER_FONTTYPE = 1

Global Const SHIFTJIS_CHARSET = 128

Global Const SYMBOL_CHARSET = 2

Global Const TA_CENTER = 6
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24
Global Const TA_LEFT = 0
Global Const TA_NOUPDATECP = 0
Global Const TA_RIGHT = 2
Global Const TA_TOP = 0
Global Const TA_UPDATECP = 1

Global Const TMPF_DEVICE = 8
Global Const TMPF_FIXED_PITCH = 1
Global Const TMPF_TRUETYPE = 4
Global Const TMPF_VECTOR = 2

Global Const TRUETYPE_FONTTYPE = 4

Global Const VARIABLE_PITCH = 2

