VERSION 5.00
Begin VB.Form TextForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text"
   ClientHeight    =   4380
   ClientLeft      =   1200
   ClientTop       =   1860
   ClientWidth     =   5190
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1695
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
   Begin VB.ComboBox Style_Combo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   1695
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   120
      Width           =   4995
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   1695
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   120
      Width           =   4995
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alignment"
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   2100
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Right"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1140
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Center"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Left"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Cancel_Btn 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3900
      TabIndex        =   8
      Top             =   3300
      Width           =   1155
   End
   Begin VB.CommandButton OK_Btn 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   420
      Left            =   3900
      TabIndex        =   7
      Top             =   3840
      Width           =   1155
   End
   Begin VB.TextBox SizeBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   780
      MaxLength       =   4
      TabIndex        =   6
      Top             =   3900
      Width           =   930
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   1740
      Max             =   1
      Min             =   999
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3900
      Value           =   12
      Width           =   240
   End
   Begin VB.ListBox FontList 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label ViewFont 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Artisan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   1860
      Width           =   4935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Size:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
End
Attribute VB_Name = "TextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Z As Integer
Dim LastFontSize As Integer
Dim LastFontName As String

Private Sub Cancel_Btn_Click()
 
 Artisan.ArtWork.FontName = LastFontName
 Artisan.ArtWork.FontSize = LastFontSize
 Artisan.ArtWork.FontBold = False
 Artisan.ArtWork.FontItalic = False

 Artisan.Txt.FontName = LastFontName
 Artisan.Txt.FontSize = LastFontSize
 Artisan.Txt.FontBold = False
 Artisan.Txt.FontItalic = False
 Artisan.Txt.Width = 15
 Align = 0

 Unload TextForm

End Sub

Private Sub CarriageReturns()

Static curindex As Long
Dim Find As String
Find = Chr(13)

Do
  curindex = InStr(curindex + 1, Text1(Z).Text, Find)
  If curindex > 0 Then
    
    Text1(Z).SelStart = curindex - 1
    Text1(Z).SelLength = Len(Find)
    CR_LF = CR_LF + 1
  End If

Loop While curindex > 0

  crCounter = CR_LF

End Sub

Private Sub FontList_Click()

  ViewFont.FontName = FontList.Text
  Artisan.ArtWork.FontName = ViewFont.FontName
  Artisan.Txt.FontName = ViewFont.FontName
  Artisan.Txt.Width = 15

End Sub

Private Sub Form_Load()

TextForm.Move (Screen.Width - TextForm.Width) / 2, (Screen.Height - TextForm.Height) / 2

 Style_Combo.AddItem "Normal"
 Style_Combo.AddItem "Bold"
 Style_Combo.AddItem "Italic"
 Style_Combo.AddItem "Bold-Italic"

 Text1(1).Visible = False
 Text1(2).Visible = False
 
 Sizebox.Text = Artisan.ArtWork.FontSize
 LastFontName = Artisan.ArtWork.FontName
 LastFontSize = Sizebox.Text
 ViewFont.FontSize = 18

For R = 0 To 20 'Printer.FontCount - 1
    FontNames$ = Printer.Fonts(R)
    FontList.AddItem FontNames$
Next R
  
For i = 0 To FontList.ListCount - 1
  If FontList.List(i) = Artisan.ArtWork.FontName Then FontList.ListIndex = i
Next i
  
  Style_Combo.ListIndex = 0

'In order to display selected text you
'would need to change StoredText(objectnumber)
'to a global declaration as well as an object
'number (B) pointer.
  
  crCounter = 0
  
End Sub

Private Sub Form_Resize()

'TextForm.Move (Screen.Width - TextForm.Width) / 2, (Screen.Height - TextForm.Height) / 2

End Sub

Private Sub OK_Btn_Click()

If Text1(Z).Text > "" Then 'print attributes

    Artisan.ArtWork.FontName = ViewFont.FontName
    Artisan.ArtWork.FontBold = ViewFont.FontBold
    Artisan.ArtWork.FontItalic = ViewFont.FontItalic

Artisan.Txt.Visible = False
Artisan.Txt.FontName = ViewFont.FontName
Artisan.Txt.FontSize = Val(Sizebox.Text)
If Zoom_In = True Then Artisan.Txt.FontSize = Artisan.Txt.FontSize * TextPoint
If Zoom_Out = True Then Artisan.Txt.FontSize = Artisan.Txt.FontSize * TextPoint

Artisan.Txt.Caption = Text1(Z).Text
Artisan.StatusLabel.Caption = ViewFont.FontName

  Select Case True
      Case Option1(0)
           Align = 0 'left
      Case Option1(1)
           Align = 6 'center
      Case Option1(2)
           Align = 2 'right
  End Select

CarriageReturns

End If

Unload TextForm

End Sub

Private Sub Option1_Click(Index As Integer)

Select Case Index
    Case 0 'Align Left
Text1(0).Text = Text1(Z).Text
Text1(0).Visible = True: Z = Index: Text1(0).SetFocus
Text1(1).Visible = False: Text1(2).Visible = False
    Case 1 'Align Centre
Text1(1).Text = Text1(Z).Text
Text1(1).Visible = True: Z = Index: Text1(1).SetFocus
Text1(0).Visible = False: Text1(2).Visible = False
    Case 2 'Align Right
Text1(2).Text = Text1(Z).Text
Text1(2).Visible = True: Z = Index: Text1(2).SetFocus
Text1(0).Visible = False: Text1(1).Visible = False

End Select

End Sub

Private Sub SizeBox_Change()

On Error GoTo BadInput

  Pts = Val(Sizebox.Text)
  Pt2 = Fix(Pts)
  VScroll1.value = Pt2: Exit Sub

BadInput:
  Sizebox.Text = "12"
  Exit Sub

End Sub

Private Sub Style_Combo_Click()

Select Case Style_Combo.Text
  Case "Normal"
    ViewFont.FontBold = False: ViewFont.FontItalic = False
  Case "Bold"
    ViewFont.FontBold = True: ViewFont.FontItalic = False
  Case "Italic"
    ViewFont.FontItalic = True: ViewFont.FontBold = False
  Case "Bold-Italic"
    ViewFont.FontBold = True: ViewFont.FontItalic = True
End Select
 
End Sub

Private Sub Text1_Change(Index As Integer)

  Z = Index: ViewFont.Caption = Left$(Text1(Z).Text, 10)

End Sub

Private Sub VScroll1_Change()

  Sizebox.Text = Format$(VScroll1.value)

  Artisan.ArtWork.FontSize = Sizebox.Text

  Artisan.Txt.Visible = False
  Artisan.Txt.FontSize = Sizebox.Text
  Artisan.Txt.Width = 15
  Artisan.Txt.Visible = True

End Sub

