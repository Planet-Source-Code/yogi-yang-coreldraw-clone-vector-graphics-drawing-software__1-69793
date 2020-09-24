Attribute VB_Name = "ART3_7"
DefInt A-Z

'Text Labels
Global Align As Integer
Global TextPoint As Single
Global CarriageCount(1000) As Integer
Global crCounter As Integer

'Zoom Bar Flags
Global Zoom_In, Zoom_Out  As Integer

' Bezier Curves
Global Convert_Curve%
Global Convert_Line%
Global FlagPrinter%

'Type POINTS
' x As Integer
' y As Integer
'End Type
Type POINTS
    x As Long
    y As Long
End Type

'Type Rect
'    Left As Integer
'    Top As Integer
'    Right As Integer
'    Bottom As Integer
'End Type
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

''''
'Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI)
'Declare Function GetMenu Lib "user32" (ByVal hwnd)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declare Function GetSubMenu Lib "User32" (ByVal hWnd, ByVal Position)
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Declare Function SendMessageByNum% Lib "user32" Alias "PostMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu, ByVal E2, ByVal X, ByVal Y, ByVal r2, ByVal hWnd, ByVal r3&)
'Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" (ByVal Appname$, ByVal KeyName$, ByVal NewString$, ByVal FileName$)
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''''''''''''''''
'odds & ends
Global Const SRCCOPY = &HCC0020
Global Const SRCAND = &H8800C6
Global Const SRCINVERT = &H660046
Global Const Blackness = &H42&
Global Const MF_BYPOSITION = &H400
Global Const BITSPIXEL = 12
Global Const HELP_CONTEXT = &H1
Global Const HELP_QUIT = &H2
Global Const HELP_HELPONHELP = &H4

'KeyWord Names
Global Const KEY_NAME = "ARTAPI"
Global Const KEY_ENTRY = "Colors Row"
Global Const INI_FILENAME = "ARTAPI.INI"

'Default Color palette values
Global Const BAS_PALETTE1 = "16777215 14737632 12632319 12640511 14745599 12648384 16777152 16761024 16761087 00000192 00016576 00049344 00049152 12632064 12582912 12583104"
Global Const BAS_PALETTE2 = "12632256 04210752 08421631 08438015 08454143 08454016 16777088 16744576 16744703 00000128 00016512 00032896 00032768 08421376 08388608 08388736"
Global Const BAS_PALETTE3 = "08421504 00000000 00000255 00033023 00065535 00065280 16776960 16711680 16711935 00000064 04210816 00016448 00016384 04210688 04194304 04194368"

'Default value for ARTAPI Filename field
Global Const UNTITLED = "[Untitled]"

'ARTAPI's ToolPalette ID's
Global Const PointerTool = 0
Global Const BezierTool = 1
Global Const ZoomTool = 2
Global Const LineTool = 3
Global Const BoxTool = 4
Global Const CircleTool = 5
Global Const OutLineTool = 6
Global Const PaintTool = 7
Global Const TextTool = 8
Global Const Poly_line = 9

'ARTAPI's ZoomBar  ID's
Global Const ZoomIn = 0
Global Const ZoomOut = 1
Global Const ZoomAll = 2
Global Const ZoomPage = 3

'ARTAPI's ToolLine ID's
Global Const LineColor = 0
Global Const NoLine = 1
Global Const HairLine = 2
Global Const OnePt = 3
Global Const TwoPt = 4
Global Const FourPt = 5
Global Const EightPt = 6
Global Const TwelvePt = 7

'ARTAPI ToolFill ID's
Global Const Palette = 0
Global Const NoFill = 1
Global Const WhiteBrush = 2
Global Const Brush10 = 3
Global Const Brush20 = 4
Global Const Brush50 = 5
Global Const Brush70 = 6
Global Const BlackBrush = 7

'Index's into the ColorPalettes' Scrl_RGB()
Global Const RED_ELEMENT = 0
Global Const GREEN_ELEMENT = 1
Global Const BLUE_ELEMENT = 2

' Color palette array
Global Colors(47)  As Long
Global PenColor As Long
Global BrushColor As Long

'ARTAPI Global Variables
Global Opaque%
Global CurrentTool%
Global LastToolUsed%
Global CurrentLine%
Global CurrentBrush%
Global CurrentZoom%
Global ColorIndex%
Global ColorString     As String
' Miscelaneous
Global Text              As String
Global CRLF              As String * 2
Global A_TAB             As String * 1
' WindowState
Global Const MINIMIZED = 1
' Clipboard formats
Global Const CF_BITMAP = 2
' Show parameters
Global Const MODAL = 1
Global Const MODELESS = 0
' Colors
Global Const BLACK = &H0&
Global Const DARKBLUE = &H800000
Global Const DARKGREEN = &H8000&
Global Const DARKCYAN = &H808000
Global Const DARKRED = &H80&
Global Const DARKPURPLE = &H800080
Global Const DARKYELLOW = &H8080&
Global Const DARKGRAY = &H808080
Global Const GRAY = &HC0C0C0
Global Const BLUE = &HFF0000
Global Const GREEN = &HFF00&
Global Const CYAN = &HFFFF00
Global Const RED = &HFF&
Global Const PURPLE = &HFF00FF
Global Const YELLOW = &HFFFF&
Global Const WHITE = &HFFFFFF
' MousePointer
Global Const DEFAULT = 0
Global Const CROSSHAIR = 2
Global Const SIZE = 5
Global Const UP_ARROW = 10
Global Const HOURGLASS = 11
' DrawMode Values
Global Const Black_Pen = 1
Global Const INVERSE = 6
Global Const Xor_Pen = 7
Global Const COPY_PEN = 13
' DrawStyle property Values
Global Const SOLID = 0
Global Const DOT = 2
' FillStyle Values
Global Const TRANSPARENT = 1
' Key Codes
Global Const KEY_BACK = &H8
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_SPACE = &H20
Global Const KEY_F1 = &H70
' Shift parameter masks
Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4
' Button parameter masks
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4
' ScaleModes
Global Const TWIPS = 1
Global Const PIXELS = 3
' MsgBox return values
Global Const MBYES = 6
Global Const MBNO = 7
' PageSize
Global PageX1 As Integer
Global PageY1 As Integer
Global PageX2 As Integer
Global PageY2 As Integer

Global POINTAPI As POINTS

Sub Display_Color_Palette(ColorBar As Control)
    
    ColorBar.Scale (0, 0)-(48, 1)

    For i = 0 To 47
        ColorBar.Line (i, 0)-(i + 1, 1), Colors(i), BF
        If i Then ColorBar.Line (i, 0)-(i, 3)
    Next i

End Sub

Sub Update_Mouse_Colors(Button, x As Single, y As Single)

Dim color As Long
    
  ColorIndex = Fix(x) + Fix(y) * 47
  color = Colors(ColorIndex)

If Button = RIGHT_BUTTON Or CurrentTool = 6 Then
  PenColor = color
  Artisan.StatusColor.BackColor = color
  Exit Sub
End If

  BrushColor = color
  Artisan.StatusPaint.BackColor = color

End Sub

