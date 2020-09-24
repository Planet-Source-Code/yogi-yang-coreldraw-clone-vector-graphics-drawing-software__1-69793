VERSION 5.00
Begin VB.Form ColorPalette 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Palette"
   ClientHeight    =   4125
   ClientLeft      =   1410
   ClientTop       =   2145
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ART3_3.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   4965
   Begin VB.CommandButton OKbtn 
      Appearance      =   0  'Flat
      Caption         =   "Cancel Edit "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   15
      Top             =   2340
      Width           =   1875
   End
   Begin VB.CommandButton OKbtn 
      Appearance      =   0  'Flat
      Caption         =   "Save Edit "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   1920
      Width           =   1875
   End
   Begin VB.CommandButton Cmd_Reset 
      Appearance      =   0  'Flat
      Caption         =   "&Reset Slides"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2340
      Width           =   2655
   End
   Begin VB.PictureBox PaletteBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   4725
      TabIndex        =   0
      Top             =   2760
      Width           =   4755
   End
   Begin VB.PictureBox ColorBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   4710
      TabIndex        =   12
      Top             =   1500
      Width           =   4740
   End
   Begin VB.HScrollBar RGBscroll 
      Height          =   240
      Index           =   2
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   6
      Top             =   1200
      Value           =   1
      Width           =   1950
   End
   Begin VB.TextBox RGB_TextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   2160
      TabIndex        =   9
      Top             =   1140
      Width           =   615
   End
   Begin VB.HScrollBar RGBscroll 
      Height          =   240
      Index           =   1
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   5
      Top             =   720
      Width           =   1950
   End
   Begin VB.TextBox RGB_TextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   660
      Width           =   615
   End
   Begin VB.TextBox RGB_TextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   180
      Width           =   615
   End
   Begin VB.HScrollBar RGBscroll 
      Height          =   240
      Index           =   0
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox DragColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "ART3_3.frx":030A
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   3300
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label ColorName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Color Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Label Lbl_Blue 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Lbl_Green 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Green"
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Lbl_Red 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "ColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim PaletteColors(47) As Long

Private Sub Cmd_Reset_Click()

    GetRGBcolors

End Sub

Private Sub ColorBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim color As Long
X = Fix(X): Y = Fix(Y)

ColorIndex = X + Y * 47

If X >= 0 And X < 48 And Y >= 0 And Y < 2 Then
   
  color = Colors(ColorIndex)
  GetRGBcolors
  ColorName.Caption = "Color Name: " + Format$(DragColor.BackColor)

End If

End Sub

Private Sub ColorBar_Paint()

    Display_Color_Palette ColorBar

End Sub

Private Sub DragColor_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

  If State = 0 Then DragColor.Cls
  DragColor.ForeColor = (QBColor(15 * Rnd))
  DragColor.CurrentX = 6
  DragColor.Print "DRAG COLOR"
  If DragColor.CurrentY > 100 Then DragColor.CurrentY = 0

End Sub

Private Sub Form_Load()
    
ColorPalette.Left = (Screen.Width - Width) / 2
ColorPalette.Top = (Screen.Height - Height) / 2

  BrushColor = Artisan.StatusPaint.BackColor
  RGBscroll(0).value = BrushColor And &HFF&
  RGBscroll(1).value = (BrushColor And &HFF00&) \ 256
  RGBscroll(2).value = (BrushColor And &HFF0000) \ 65536

  For i = 0 To 47
    PaletteColors(i) = Colors(i)
  Next i

End Sub

Private Sub Form_Resize()

ColorPalette.Left = (Screen.Width - ColorPalette.Width) / 2
ColorPalette.Top = (Screen.Height - ColorPalette.Height) / 2

End Sub

Private Sub GetRGBcolors()

  R1 = (Colors(ColorIndex) And &HFF)
  G2 = (Colors(ColorIndex) And &HFF00&) \ 256
  B3 = (Colors(ColorIndex) And &HFF0000) \ 65536

'R1 = Colors(ColorIndex) Mod 256        'same as above
'G2 = Colors(ColorIndex) \ 256 Mod 256
'B3 = Colors(ColorIndex) \ 65536 Mod 256

'R1 = Colors(ColorIndex) And &HFF&
'G2 = (Colors(ColorIndex) \ 2 ^ 8) And &HFF&
'B3 = (Colors(ColorIndex) \ 2 ^ 16) And &HFF&

  RGBscroll(RED_ELEMENT).value = R1
  RGBscroll(GREEN_ELEMENT).value = G2
  RGBscroll(BLUE_ELEMENT).value = B3

End Sub

Private Sub OKbtn_Click(Index As Integer)

  If Index = 0 Then
    For i = 0 To 47
      Colors(i) = PaletteColors(i)
    Next i
  End If

  Unload ColorPalette
  
End Sub

Private Sub PaletteBox_DragDrop(Source As Control, X As Single, Y As Single)

X = Fix(X) 'will return the location
If Fix(Y) = 1 Then X = Fix(X) + 15
If Fix(Y) = 2 Then X = Fix(X) + 32

    PaletteColors(X) = Val(DragColor.BackColor)
    PaletteBox_Paint 'Update Color in Palette Box
    
End Sub

Private Sub PaletteBox_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

  If State = 0 Then DragColor.Cls
  DragColor.ForeColor = (QBColor(15 * Rnd))
  DragColor.CurrentX = 6
  DragColor.Print "DROP COLOR"
  If DragColor.CurrentY > 100 Then DragColor.CurrentY = 0

End Sub

Private Sub PaletteBox_Paint()

  PaletteBox.Scale (0, 0)-(16, 3)

  For R = 1 To 3

    For i = 0 To 15
        PaletteBox.Line (i, R - 1)-(i + 1, R), PaletteColors(i + NR), BF
        PaletteBox.Line (i, R - 1)-(i + 1, R), QBColor(0), B
    Next i
   NR = NR + 16 'Next row
 Next R

End Sub

Private Sub RGB_TextBox_Change(Index As Integer)
    
On Error GoTo BadInput
  
  RGBscroll(Index).value = RGB_TextBox(Index).Text: Exit Sub

BadInput:
  RGB_TextBox(Index).Text = RGBscroll(Index).value: Exit Sub
    
End Sub

Private Sub RGB_TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
    
  If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then
     KeyAscii = 0: Beep
  End If

End Sub

Private Sub RGBscroll_Change(Index As Integer)

  RGBscroll_scroll Index

End Sub

Private Sub RGBscroll_scroll(Index As Integer)

DragColor.BackColor = RGB(RGBscroll(0).value, RGBscroll(1).value, RGBscroll(2).value)

   For i = 0 To 2
       RGB_TextBox(i).Text = Format$(RGBscroll(i).value)
   Next i

End Sub

