VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Artisan 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Artisan - Corel Draw Clone"
   ClientHeight    =   6435
   ClientLeft      =   1680
   ClientTop       =   1755
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ART3_0.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   3360
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox RulerCross 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
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
      Height          =   345
      Left            =   540
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox StatusArea 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4980
      Width           =   7170
      Begin VB.PictureBox ColorMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   475
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   900
         Width           =   7155
         Begin VB.PictureBox ColorBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ClipControls    =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   90
            Picture         =   "ART3_0.frx":030A
            ScaleHeight     =   1
            ScaleMode       =   0  'User
            ScaleWidth      =   48
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   90
            Width           =   6930
         End
      End
      Begin VB.PictureBox StatusColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   390
         Left            =   6540
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   390
         Width           =   450
         Begin VB.PictureBox StatusPaint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ClipControls    =   0   'False
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
            Height          =   315
            Left            =   60
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   60
            Width           =   345
         End
      End
      Begin VB.HScrollBar ColorScroll 
         Height          =   255
         LargeChange     =   80
         Left            =   45
         Max             =   2000
         SmallChange     =   16
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   315
         Width           =   510
      End
      Begin VB.CommandButton No_Fill 
         Appearance      =   0  'Flat
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   585
         Width           =   495
      End
      Begin VB.HScrollBar HScroll 
         Height          =   240
         LargeChange     =   100
         Left            =   810
         Max             =   200
         Min             =   -200
         SmallChange     =   10
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   45
         Width           =   5985
      End
      Begin VB.CommandButton TestButton 
         Appearance      =   0  'Flat
         Caption         =   "test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   780
      End
      Begin VB.Label XYLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   390
         Left            =   750
         TabIndex        =   34
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label StatusLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   390
         Left            =   1935
         TabIndex        =   33
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.PictureBox Ruler 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   4335
      Index           =   0
      Left            =   540
      ScaleHeight     =   4305
      ScaleWidth      =   300
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   405
      Width           =   330
   End
   Begin VB.VScrollBar VScroll 
      Height          =   4725
      LargeChange     =   100
      Left            =   6930
      Max             =   250
      Min             =   -250
      SmallChange     =   10
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Ruler 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   345
      Index           =   1
      Left            =   900
      ScaleHeight     =   315
      ScaleWidth      =   5985
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
   End
   Begin VB.PictureBox PaintButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1035
      Picture         =   "ART3_0.frx":0750
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox ZoomBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1020
      Picture         =   "ART3_0.frx":18D2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   126
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox LineButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1035
      Picture         =   "ART3_0.frx":21D4
      ScaleHeight     =   34
      ScaleMode       =   0  'User
      ScaleWidth      =   256
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox PictureMask 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   945
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   391
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5865
      Begin VB.PictureBox ArtWork 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   2100
         ScaleHeight     =   2295
         ScaleWidth      =   3540
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   3570
         Begin VB.PictureBox ScaleBox 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   330
            Left            =   900
            ScaleHeight     =   300
            ScaleWidth      =   300
            TabIndex        =   38
            Top             =   990
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   0
            Left            =   630
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox BezNode2 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   105
            Left            =   45
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   855
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox BezNode1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   105
            Left            =   45
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   675
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox BezHandle2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   105
            Left            =   45
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   495
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox BezHandle1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   105
            Left            =   45
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   2
            Left            =   1980
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1710
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   3
            Left            =   630
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1710
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   6
            Left            =   1320
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox Node 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   90
            Index           =   2
            Left            =   1845
            ScaleHeight     =   90
            ScaleWidth      =   90
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1530
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox Node 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   90
            Index           =   3
            Left            =   870
            ScaleHeight     =   90
            ScaleWidth      =   90
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1575
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   7
            Left            =   630
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   5
            Left            =   1980
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox Node 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   90
            Index           =   1
            Left            =   1800
            ScaleHeight     =   90
            ScaleWidth      =   90
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox Node 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   90
            Index           =   0
            Left            =   855
            ScaleHeight     =   90
            ScaleWidth      =   90
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   1
            Left            =   1980
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.PictureBox Handle 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
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
            Height          =   120
            Index           =   4
            Left            =   1305
            MousePointer    =   2  'Cross
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Txt 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   1440
            TabIndex        =   22
            Top             =   900
            Visible         =   0   'False
            Width           =   15
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   105
            Left            =   45
            Top             =   1305
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   105
            Left            =   45
            Top             =   1080
            Visible         =   0   'False
            Width           =   105
         End
      End
   End
   Begin VB.PictureBox ToolButtons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   0
      Picture         =   "ART3_0.frx":32D6
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   540
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   +{F12}
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "Save &As..."
         Index           =   3
         Shortcut        =   {F12}
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&Print"
         Index           =   4
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Delete"
         Index           =   6
      End
   End
   Begin VB.Menu Menu_Arrange 
      Caption         =   "&Arrange"
      Begin VB.Menu Menu_ArrangeSelection 
         Caption         =   "To &Front"
         Index           =   0
      End
      Begin VB.Menu Menu_ArrangeSelection 
         Caption         =   "To &Back"
         Index           =   1
      End
      Begin VB.Menu Menu_ArrangeSelection 
         Caption         =   "POPUP MENU"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu Menu_Color 
      Caption         =   "&Color"
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "&Color Palette"
         Index           =   0
         Begin VB.Menu Menu_Palette 
            Caption         =   "&Edit Palette"
            Index           =   0
         End
         Begin VB.Menu Menu_Palette 
            Caption         =   "&Default Palette"
            Index           =   1
         End
         Begin VB.Menu Menu_Palette 
            Caption         =   "&Custom Palette"
            Index           =   2
         End
         Begin VB.Menu Menu_Palette 
            Caption         =   "Save As Custom"
            Index           =   3
         End
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "Show &Rulers"
         Checked         =   -1  'True
         Index           =   2
      End
   End
End
Attribute VB_Name = "Artisan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Const Limit = 1000
Dim Draw As Integer
Dim FillValue(Limit) As Integer 'No Fill
Dim LineValue(Limit) As Integer 'No Line
Dim OFFSETx, OFFSETy, PtInRectBasic As Integer

Dim WIDE, HIGH As Integer 'Monitor Type
Dim PixelX, PixelY As Integer

Dim SSM As Single 'Side Scroll Move
Dim BSM As Single 'Bottom Scroll Move

Dim LastMove As Integer
Dim LastMenu As Integer
Dim FlagMove As Integer

Dim V As Integer 'Ruler Guide Lines
Dim TGL As Integer
Dim NewGuideLine As Integer
Dim GuideX(Limit) As Long
Dim GuideY(Limit) As Long

Dim Zoom  As Integer 'Zoom Offsets
Dim NodeZoom As Integer
Dim Scaling As Integer
Dim OldSW As Double 'Scales
Dim OldSH As Double
Dim OldRH As Single 'Rulers
Dim OldRW As Single
Dim PerSW As Single 'Percentage of scales
Dim PerSH As Single

Dim Xstr As String, Ystr As String
Dim OutLineColor(Limit) As Long
Dim PaintColor(Limit) As Long
Dim LineWidth(Limit) As Integer

Dim TextX, TextY  As Integer  ' Text Values
Dim FlagAlign(Limit) As Integer
Dim OFFSET(Limit) As Integer
Dim TextSize(Limit) As Single
Dim TextFont(Limit)  As String
Dim TextBold(Limit) As Integer
Dim TextItalic(Limit) As Integer
Dim Storedtext(Limit) As String

Dim B As Integer  ' Current Object (Bounding Box)
Dim Total As Integer  ' Total Number of Objects

Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long

Dim XX1 As Long
Dim YY1 As Long
Dim XX2 As Long
Dim YY2 As Long
Dim W(Limit) As Long 'West or Left
Dim E(Limit) As Long 'East or Right
Dim N(Limit) As Long 'North or Top
Dim S(Limit) As Long 'South or Bottom

Dim FlagBezier(Limit) As Integer
Dim FlagLine(Limit) As Integer
Dim FlagBox(Limit) As Integer
Dim FlagCircle(Limit) As Integer
Dim FlagText(Limit)  As Integer

Dim Radius As Integer
Dim XCenter As Single
Dim YCenter As Single
Dim Aspect As Single

Dim SideRuler As Integer
Dim TopRuler As Integer
Dim X1Ruler As Integer
Dim Y1Ruler As Integer
Dim X2Ruler As Integer
Dim Y2Ruler As Integer

Dim DeleteObject As Integer 'DELETE

'Bezier Curve
Private Declare Function Polyline Lib "GDI" (ByVal hDC%, p As POINTS, ByVal NUM_PTS%) As Integer
Const MAX_POINTS = 33
Dim P_Index(1 To MAX_POINTS) As POINTS
Dim BezErase As Integer
Dim ShowDotted As Integer
Dim p As Integer

Dim p1 As POINTS 'Control Points / Handles
Dim h1 As POINTS
Dim h2 As POINTS
Dim p2 As POINTS

Dim P1X(Limit) As Long
Dim P1Y(Limit) As Long
Dim P2X(Limit) As Long
Dim p2Y(Limit) As Long
Dim h1X(Limit) As Long
Dim h1Y(Limit) As Long
Dim h2X(Limit) As Long
Dim h2Y(Limit) As Long

Dim C1(1 To MAX_POINTS) As POINTS
Dim D1(1 To MAX_POINTS) As POINTS 'Curve Points
Dim D2(1 To MAX_POINTS) As POINTS
Dim D3(1 To MAX_POINTS) As POINTS
Dim E1(1 To MAX_POINTS) As POINTS
Dim E2(1 To MAX_POINTS) As POINTS
Const YELLOW = &HFFFF&

Private Sub Artwork_DblClick()

If FlagLine(B) = B Then NodeEdit!To_Line.Enabled = False: NodeEdit!To_Curve.Enabled = True
If FlagBezier(B) = B Then NodeEdit!To_Curve.Enabled = False: NodeEdit!To_Line.Enabled = True
If CurrentTool = 1 Then NodeEdit.Show MODAL Else Exit Sub

Select Case True
   Case Convert_Curve
      FlagLine(B) = 0: FlagBezier(B) = B
      BezErase = True: BezHandles
      BezHandle1.Visible = True: BezHandle2.Visible = True
      BezNode1.Visible = True: BezNode2.Visible = True
      Convert_Curve = False: NodeEdit!To_Line.Enabled = True
   Case Convert_Line
      FlagBezier(B) = 0:  FlagLine(B) = B
      BezNode1.Visible = False: BezNode2.Visible = False
      W(B) = P1X(B): N(B) = P1Y(B)
      E(B) = P2X(B): S(B) = p2Y(B)
      Nodes
      Convert_Line = False: NodeEdit!To_Curve.Enabled = True
      UpdateObjects
End Select

End Sub

Private Sub ArtWork_KeyPress(KeyAscii As Integer)

If CurrentTool = TextTool Then

Select Case KeyAscii
   Case 13 'Return
     crCounter = crCounter + 1
   Case 8  'BackSpace
     If Txt.Caption = "" Then Exit Sub
     If Txt.Caption > "" Then
     If Mid$(Txt.Caption, Len(Txt.Caption)) = Chr(13) Then crCounter = crCounter - 1
     Txt.Caption = Mid$(Txt.Caption, 1, Len(Txt.Caption) - 1): Exit Sub
     End If
End Select

  Txt.Caption = Txt.Caption + Chr(KeyAscii)
  StatusLabel.Caption = Txt.FontName
  If Txt.Caption > "" Then Txt.BorderStyle = 0

End If

End Sub

Private Sub ArtWork_LostFocus()

  If CurrentTool = TextTool And Txt.Caption <> "" Then ReDrawText

End Sub

Private Sub ArtWork_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

ZoomBar.Visible = False
LineButtons.Visible = False
PaintButtons.Visible = False
If SideRuler = True Or TopRuler = True Then Exit Sub
  
If Button = LEFT_BUTTON Then N_H_Clear

Select Case CurrentTool
    Case PointerTool
      Find_Object X, Y
      BoundingBox_Down X, Y
    Case BezierTool
      If BezHandle1.Visible = True Then ShowDotted = False: BezUpdate
      BoundingBox_Down X, Y
    Case LineTool To CircleTool
      BoundingBox_Down X, Y
    Case TextTool
      If Txt.Caption > "" Then Align = 0: ReDrawText
      If Button = LEFT_BUTTON Then
        Txt.BorderStyle = 0
        Dim tm As TEXTMETRIC
        API = GetTextMetrics(ArtWork.hDC, tm)
        Leading = (tm.tmExternalLeading + tm.tmInternalLeading) * PixelY
        Txt.Left = X: Txt.Top = Y - (Txt.Height - Leading)
      End If
End Select

If CurrentTool <> 1 Then BezNode1.Visible = False: BezNode2.Visible = False

End Sub

Private Sub Artwork_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Xstr = Format$(X \ PixelX, "000")
Ystr = Format$(Y \ PixelY, "000")
XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"

RulerGuides Index, Button, Shift, X, Y

If SideRuler = True Or TopRuler = True Then
  GuideLines Index, Button, X, Y: Exit Sub
End If

If Button = LEFT_BUTTON Then
   BoundingBox_Move Shift, X, Y
    X2 = X: Y2 = Y
   BoundingBox_Move Shift, X, Y
End If

End Sub

Private Sub ArtWork_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If CurrentTool <> 8 Then ArtWork.MousePointer = 0

If SideRuler = True Or TopRuler = True Then
  Ruler_MouseUp Index, Button, Shift, X, Y
  Exit Sub
End If

If Button = LEFT_BUTTON Then
Select Case CurrentTool
    Case PointerTool
      If FlagMove = True Then UpdateObjects: FlagMove = False: Exit Sub
      If Node(0).Visible Then Handles: Exit Sub
      If BezNode1.Visible Then Handles: Exit Sub
      If MoveMouse() = True Then
        ArtWork.Line (X1, Y1)-(X2, Y2), YELLOW, B
        FindAllObjects X1, Y1, X2, Y2, X, Y
      End If
    Case BezierTool
      If MoveMouse() = True Then
        ArtWork.Line (X1, Y1)-(X2, Y2), YELLOW, B
      End If
    Case ZoomTool
      '''''''''''''''''''''''
    Case LineTool To CircleTool
      If MoveMouse() = True Then
        BoundingBox_Up X1, Y1, X2, Y2
        If B = Total Then Draw = True: UpdateObjects
      End If
    Case TextTool
      Txt.BorderStyle = 1: Txt.Width = 15: Txt.Visible = True

End Select

ArtWork.DrawWidth = 1: ArtWork.DrawStyle = SOLID

End If

If Button = 2 And Menu_ArrangeSelection(3).Checked And CurrentTool < 8 Then
   Dim X_Y As POINTAPI
   GetCursorPos X_Y
R = TrackPopupMenu(GetSubMenu(GetMenu(hwnd), LastMenu), 0, X_Y.X, X_Y.Y, 0, hwnd, 0&)
ElseIf CurrentTool = TextTool And Button = 2 Then
   ArtWork.CurrentY = Txt.Top
   TextForm.Show MODAL
   If Txt.Caption > "" Then ReDrawText
End If

End Sub

Private Sub BezDraw()
'''Leftside
p = 17: BezierDivide p1, h1, h2, p2
p = 13: BezierDivide p1, D1(17), E1(17), C1(17)
p = 9: BezierDivide p1, D1(13), E1(13), C1(13)
p = 5: BezierDivide p1, D1(9), E1(9), C1(9)
p = 3: BezierDivide p1, D1(5), E1(5), C1(5)
p = 2: BezierDivide p1, D1(3), E1(3), C1(3)
'''''
p = 4: BezierDivide C1(3), E2(3), D3(3), C1(5)
p = 7: BezierDivide C1(5), E2(5), D3(5), C1(9)
p = 6: BezierDivide C1(7), D1(7), E1(7), C1(5)
p = 8: BezierDivide C1(7), E2(7), D3(7), C1(9)
p = 11: BezierDivide C1(9), E2(9), D3(9), C1(13)
p = 10: BezierDivide C1(11), D1(11), E1(11), C1(9)
p = 12: BezierDivide C1(11), E2(11), D3(11), C1(13)
p = 15: BezierDivide C1(13), E2(13), D3(13), C1(17)
p = 14: BezierDivide C1(15), D1(15), E1(15), C1(13)
p = 16: BezierDivide C1(15), E2(15), D3(15), C1(17)
'''Rightside
p = 21: BezierDivide C1(17), E2(17), D3(17), p2
p = 25: BezierDivide C1(21), E2(21), D3(21), p2
p = 29: BezierDivide C1(25), E2(25), D3(25), p2
p = 31: BezierDivide C1(29), E2(29), D3(29), p2
p = 32: BezierDivide C1(31), E2(31), D3(31), p2
'''''
p = 30: BezierDivide C1(29), D1(31), E1(31), C1(31)
p = 27: BezierDivide C1(25), D1(29), E1(29), C1(29)
p = 28: BezierDivide C1(29), E2(27), D3(27), C1(27)
p = 26: BezierDivide C1(25), D1(27), E1(27), C1(27)
p = 23: BezierDivide C1(21), D1(25), E1(25), C1(25)
p = 24: BezierDivide C1(25), E2(23), D3(23), C1(23)
p = 22: BezierDivide C1(21), D1(23), E1(23), C1(23)
p = 19: BezierDivide C1(17), D1(21), E1(21), C1(21)
p = 20: BezierDivide C1(21), E2(19), D3(19), C1(19)
p = 18: BezierDivide C1(17), D1(19), E1(19), C1(19)

If FlagPrinter = 2 Then GoTo PrintToPage

If ArtWork.AutoRedraw = False Then ArtWork.DrawMode = INVERSE
ArtWork.DrawStyle = SOLID
''''''''''''''''''''''''''''''''''''''''''''''''''''
    'P_Index(1).X = p1.X \ PIXELX
    'P_Index(1).Y = p1.Y \ PIXELY
'For P = 2 To 32
    'P_Index(P).X = C1(P).X \ PIXELX
    'P_Index(P).Y = C1(P).Y \ PIXELY
    'Next P
'P_Index(33).X = p2.X \ PIXELX
'P_Index(33).Y = p2.Y \ PIXELY
'NUM_PTS = 33
'R = Polyline(ArtWork.hDC, P_Index(1), NUM_PTS)
''''''''''''''''''''''''''''''''''''''''''''''''''''
ArtWork.PSet (p1.X, p1.Y)
For p = 2 To 32
ArtWork.Line -(C1(p).X, C1(p).Y)
Next p
ArtWork.Line -(p2.X, p2.Y)
Exit Sub

PrintToPage:

Printer.PSet (p1.X, p1.Y)
For p = 2 To 32
Printer.Line -(C1(p).X, C1(p).Y)
Next p
Printer.Line -(p2.X, p2.Y)
Exit Sub

End Sub

Private Sub BezHandle1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

XX2 = E(B): YY2 = S(B): XX1 = W(B): YY1 = N(B)

    BezHandle1.Visible = False
    BezNode1.BackColor = BLUE: BezNode2.BackColor = BLACK
    BezMove

BezDraw

End Sub

Private Sub BezHandle1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

    Xstr = Format$(X \ PixelX, "000")
    Ystr = Format$(Y \ PixelY, "000")
    XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"

If Button Then
BezDraw
ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
  h1.X = X
  h1.Y = Y
Shape1.Left = h1.X - OFBN
Shape1.Top = h1.Y - OFBN
Shape1.Visible = True
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
BezDraw
End If

End Sub

Private Sub BezHandle1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

Shape1.Visible = False
BezHandle1.Move X - OFBN, Y - OFBN
BezHandle1.ScaleLeft = BezHandle1.Left
BezHandle1.ScaleTop = BezHandle1.Top
BezHandle1.Visible = True

ArtWork.DrawStyle = 0: ArtWork.DrawWidth = 1

P1X(B) = p1.X: P1Y(B) = p1.Y
P2X(B) = p2.X: p2Y(B) = p2.Y
h1X(B) = h1.X: h1Y(B) = h1.Y
h2X(B) = h2.X: h2Y(B) = h2.Y
BezierStore
UpdateObjects
p1.X = P1X(B): p1.Y = P1Y(B): p2.X = P2X(B): p2.Y = p2Y(B)
h1.X = h1X(B): h1.Y = h1Y(B): h2.X = h2X(B): h2.Y = h2Y(B)

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

End Sub

Private Sub BezHandle2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

XX2 = E(B): YY2 = S(B): XX1 = W(B): YY1 = N(B)

BezHandle2.Visible = False
BezNode2.BackColor = BLUE: BezNode1.BackColor = BLACK
BezMove

BezDraw

End Sub

Private Sub BezHandle2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

Xstr = Format$(X \ PixelX, "000")
Ystr = Format$(Y \ PixelY, "000")
XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"

If Button Then
BezDraw
ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
  h2.X = X
  h2.Y = Y
Shape2.Left = h2.X - OFBN
Shape2.Top = h2.Y - OFBN
Shape2.Visible = True
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
BezDraw
End If

End Sub

Private Sub BezHandle2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

Shape2.Visible = False
BezHandle2.Move X - OFBN, Y - OFBN
BezHandle2.ScaleLeft = BezHandle2.Left
BezHandle2.ScaleTop = BezHandle2.Top
BezHandle2.Visible = True

ArtWork.DrawStyle = 0: ArtWork.DrawWidth = 1

P1X(B) = p1.X: P1Y(B) = p1.Y
P2X(B) = p2.X: p2Y(B) = p2.Y
h1X(B) = h1.X: h1Y(B) = h1.Y
h2X(B) = h2.X: h2Y(B) = h2.Y
BezierStore
UpdateObjects
p1.X = P1X(B): p1.Y = P1Y(B): p2.X = P2X(B): p2.Y = p2Y(B)
h1.X = h1X(B): h1.Y = h1Y(B): h2.X = h2X(B): h2.Y = h2Y(B)

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

End Sub

Private Sub BezHandles()

OFBN = 45 + NodeZoom

 BezNode1.Move W(B) - OFBN, N(B) - OFBN
 BezNode2.Move E(B) - OFBN, S(B) - OFBN

 p1.X = W(B)
 p1.Y = N(B)
 p2.X = E(B)
 p2.Y = S(B)
'''''''''''''
ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
HalfLeft = (E(B) + W(B)) \ 2
HalfTop = (S(B) + N(B)) \ 2
    BezHandle1.Left = (W(B) + HalfLeft) \ 2 - OFBN
    BezHandle1.Top = (N(B) + HalfTop) \ 2 - OFBN
    BezHandle2.Left = (E(B) + HalfLeft) \ 2 - OFBN
    BezHandle2.Top = (S(B) + HalfTop) \ 2 - OFBN
 BezHandle1.ScaleLeft = BezHandle1.Left
 BezHandle1.ScaleTop = BezHandle1.Top
 BezHandle2.ScaleLeft = BezHandle2.Left
 BezHandle2.ScaleTop = BezHandle2.Top
 BezNode1.ScaleLeft = BezNode1.Left
 BezNode1.ScaleTop = BezNode1.Top
 BezNode2.ScaleLeft = BezNode2.Left
 BezNode2.ScaleTop = BezNode2.Top
 ReScale  '''If ZoomTool Used

h1.X = BezHandle1.Left + OFBN
h1.Y = BezHandle1.Top + OFBN
h2.X = BezHandle2.Left + OFBN
h2.Y = BezHandle2.Top + OFBN

P1X(B) = p1.X: P1Y(B) = p1.Y
P2X(B) = p2.X: p2Y(B) = p2.Y
h1X(B) = h1.X: h1Y(B) = h1.Y
h2X(B) = h2.X: h2Y(B) = h2.Y

ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

End Sub

Private Sub BezierDivide(p1 As POINTS, h1 As POINTS, h2 As POINTS, p2 As POINTS)

D1(p).X = (p1.X \ 2 + h1.X \ 2)
D1(p).Y = (p1.Y \ 2 + h1.Y \ 2)
D2(p).X = (h1.X \ 2 + h2.X \ 2)
D2(p).Y = (h1.Y \ 2 + h2.Y \ 2)
D3(p).X = (p2.X \ 2 + h2.X \ 2)
D3(p).Y = (p2.Y \ 2 + h2.Y \ 2)
E1(p).X = (D1(p).X \ 2 + D2(p).X \ 2)
E1(p).Y = (D1(p).Y \ 2 + D2(p).Y \ 2)
E2(p).X = (D2(p).X \ 2 + D3(p).X \ 2)
E2(p).Y = (D2(p).Y \ 2 + D3(p).Y \ 2)
C1(p).X = (E1(p).X \ 2 + E2(p).X \ 2)
C1(p).Y = (E1(p).Y \ 2 + E2(p).Y \ 2)
'ArtWork.DrawWidth = 4
'ArtWork.PSet (C1(P).X, C1(P).Y)
'ArtWork.DrawWidth = 1

End Sub

Private Sub BezierStore()

If p1.X <= h1.X And p1.X <= h2.X And p1.X <= p2.X Then W(B) = p1.X
If h1.X <= p1.X And h1.X <= h2.X And h1.X <= p2.X Then W(B) = h1.X
If h2.X <= p1.X And h2.X <= h1.X And h2.X <= p2.X Then W(B) = h2.X
If p2.X <= p1.X And p2.X <= h1.X And p2.X <= h2.X Then W(B) = p2.X

If p1.Y <= h1.Y And p1.Y <= h2.Y And p1.Y <= p2.Y Then N(B) = p1.Y
If h1.Y <= p1.Y And h1.Y <= h2.Y And h1.Y <= p2.Y Then N(B) = h1.Y
If h2.Y <= p1.Y And h2.Y <= h1.Y And h2.Y <= p2.Y Then N(B) = h2.Y
If p2.Y <= p1.Y And p2.Y <= h1.Y And p2.Y <= h2.Y Then N(B) = p2.Y

If p1.X >= h1.X And p1.X >= h2.X And p1.X >= p2.X Then E(B) = p1.X
If h1.X >= p1.X And h1.X >= h2.X And h1.X >= p2.X Then E(B) = h1.X
If h2.X >= p1.X And h2.X >= h1.X And h2.X >= p2.X Then E(B) = h2.X
If p2.X >= p1.X And p2.X >= h1.X And p2.X >= h2.X Then E(B) = p2.X

If p1.Y >= h1.Y And p1.Y >= h2.Y And p1.Y >= p2.Y Then S(B) = p1.Y
If h1.Y >= p1.Y And h1.Y >= h2.Y And h1.Y >= p2.Y Then S(B) = h1.Y
If h2.Y >= p1.Y And h2.Y >= h1.Y And h2.Y >= p2.Y Then S(B) = h2.Y
If p2.Y >= p1.Y And p2.Y >= h1.Y And p2.Y >= h2.Y Then S(B) = p2.Y

End Sub

Private Sub BezMove()

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

If BezErase = True Then
  ArtWork.AutoRedraw = True
  ArtWork.DrawStyle = 0: ArtWork.DrawMode = 13
  ArtWork.Line (p1.X, p1.Y)-(p2.X, p2.Y), QBColor(15)
  ArtWork.AutoRedraw = False: BezErase = False
Else
  ArtWork.AutoRedraw = True
  ArtWork.DrawMode = 16
  ArtWork.DrawWidth = 1
  ArtWork.PSet (0, 0)
  BezDraw
  ArtWork.DrawWidth = 1
  ArtWork.AutoRedraw = False
End If

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

End Sub

Private Sub BezNode1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

XX2 = E(B): YY2 = S(B): XX1 = W(B): YY1 = N(B)

BezNode1.BackColor = BLUE: BezNode2.BackColor = BLACK
    
    X1 = p1.X: Y1 = p1.Y
    XX1 = X: YY1 = Y
    p = 1

End Sub

Private Sub BezNode1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

    Xstr = Format$(X \ PixelX, "000")
    Ystr = Format$(Y \ PixelY, "000")
    XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"

If Button Then
If X = XX1 And Y = YY1 Then XX1 = 0 And YY1 = 0: Exit Sub
If p = 1 Then
p = 0: BezHandle1.Visible = True: BezHandle2.Visible = True: BezMove
BezDraw
End If

BezDraw
ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
  p1.X = X
  p1.Y = Y
  h1.X = (X - X1) + h1X(B)
  h1.Y = (Y - Y1) + h1Y(B)
BezHandle1.Left = h1.X - OFBN
BezHandle1.Top = h1.Y - OFBN

Shape1.Left = p1.X - OFBN
Shape1.Top = p1.Y - OFBN
BezNode1.Visible = False: Shape1.Visible = True
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

BezDraw
End If

End Sub

Private Sub BezNode1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

If X = XX1 And Y = YY1 Then XX1 = 0 And YY1 = 0: Exit Sub

BezNode1.Move X - OFBN, Y - OFBN
BezNode1.ScaleLeft = BezNode1.Left
BezNode1.ScaleTop = BezNode1.Top
Shape1.Visible = False: BezNode1.Visible = True
BezHandle1.ScaleLeft = BezHandle1.Left
BezHandle1.ScaleTop = BezHandle1.Top

X1 = p1.X: Y1 = p1.Y

P1X(B) = p1.X: P1Y(B) = p1.Y
P2X(B) = p2.X: p2Y(B) = p2.Y
h1X(B) = h1.X: h1Y(B) = h1.Y
h2X(B) = h2.X: h2Y(B) = h2.Y
UpdateObjects
p1.X = P1X(B): p1.Y = P1Y(B): p2.X = P2X(B): p2.Y = p2Y(B)
h1.X = h1X(B): h1.Y = h1Y(B): h2.X = h2X(B): h2.Y = h2Y(B)

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
BezierStore

End Sub

Private Sub BezNode2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

XX2 = E(B): YY2 = S(B): XX1 = W(B): YY1 = N(B)

BezNode2.BackColor = BLUE: BezNode1.BackColor = BLACK
X2 = p2.X: Y2 = p2.Y
XX1 = X: YY1 = Y
p = 1

End Sub

Private Sub BezNode2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

Xstr = Format$(X \ PixelX, "000")
Ystr = Format$(Y \ PixelY, "000")
XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"

If Button Then
If X = XX1 And Y = YY1 Then XX1 = 0 And YY1 = 0: Exit Sub
If p = 1 Then
p = 0: BezHandle1.Visible = True: BezHandle2.Visible = True: BezMove
BezDraw
End If

BezDraw
ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
  p2.X = X
  p2.Y = Y
  h2.X = (X - X2) + h2X(B)
  h2.Y = (Y - Y2) + h2Y(B)
BezHandle2.Left = h2.X - OFBN
BezHandle2.Top = h2.Y - OFBN
Shape2.Left = p2.X - OFBN
Shape2.Top = p2.Y - OFBN
BezNode2.Visible = False: Shape2.Visible = True
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW

BezDraw
End If

End Sub

Private Sub BezNode2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

OFBN = 45 + NodeZoom

If X = XX1 And Y = YY1 Then XX1 = 0 And YY1 = 0: Exit Sub

BezNode2.Move X - OFBN, Y - OFBN
BezNode2.ScaleLeft = BezNode2.Left
BezNode2.ScaleTop = BezNode2.Top
Shape2.Visible = False: BezNode2.Visible = True
BezHandle2.ScaleLeft = BezHandle2.Left
BezHandle2.ScaleTop = BezHandle2.Top

X2 = p2.X: Y2 = p2.Y

P1X(B) = p1.X: P1Y(B) = p1.Y
P2X(B) = p2.X: p2Y(B) = p2.Y
h1X(B) = h1.X: h1Y(B) = h1.Y
h2X(B) = h2.X: h2Y(B) = h2.Y
UpdateObjects
p1.X = P1X(B): p1.Y = P1Y(B): p2.X = P2X(B): p2.Y = p2Y(B)
h1.X = h1X(B): h1.Y = h1Y(B): h2.X = h2X(B): h2.Y = h2Y(B)

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
BezierStore

End Sub

Private Sub BezStretch(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0, 1, 2, 4, 5, 6
     pt1 = P1Y(B): Pt2 = p2Y(B): hd1 = h1Y(B): hd2 = h2Y(B)
    Case 3, 7
     pt1 = P1X(B): Pt2 = P2X(B): hd1 = h1X(B): hd2 = h2X(B)
End Select

If pt1 = Pt2 And hd1 = hd2 Then
W(B) = XX1: N(B) = YY1: E(B) = XX2: S(B) = YY2: Exit Sub
End If

Boat1 = pt1: Boat2 = Pt2: Boat3 = hd1: Boat4 = hd2
''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case Index
    Case 0, 1, 2
      percentage! = (Y - YY1) / (YY2 - YY1): ANCHOR = YY2
    Case 3
      percentage! = (X - XX2) / (XX1 - XX2): ANCHOR = XX1
    Case 4, 5, 6
      percentage! = (Y - YY2) / (YY1 - YY2): ANCHOR = YY1
    Case 7
      percentage! = (X - XX1) / (XX2 - XX1): ANCHOR = XX2
End Select

Boat1 = ((ANCHOR - Boat1) * percentage!) + Boat1
Boat2 = ((ANCHOR - Boat2) * percentage!) + Boat2
Boat3 = ((ANCHOR - Boat3) * percentage!) + Boat3
Boat4 = ((ANCHOR - Boat4) * percentage!) + Boat4

If Index = 3 Or Index = 7 Then GoTo Xpoint Else GoTo Ypoint

Xpoint: 'Stretch from middle side Handles
p1.X = Boat1: P1X(B) = Boat1: p2.X = Boat2: P2X(B) = Boat2
h1.X = Boat3: h1X(B) = Boat3: h2.X = Boat4: h2X(B) = Boat4
BezierStore
Exit Sub

Ypoint: 'Stretch from middle top and bottom Handles
p1.Y = Boat1: P1Y(B) = Boat1: p2.Y = Boat2: p2Y(B) = Boat2
h1.Y = Boat3: h1Y(B) = Boat3: h2.Y = Boat4: h2Y(B) = Boat4
If Index = 1 Or Index = 5 Then BezierStore: Exit Sub

'''''''''''' Stretch from corner Handles '''''''''''''
pt1 = P1X(B): Pt2 = P2X(B): hd1 = h1X(B): hd2 = h2X(B)
Boat1 = pt1: Boat2 = Pt2: Boat3 = hd1: Boat4 = hd2

Select Case Index
    Case 0, 6
      percentage! = (X - XX1) / (XX2 - XX1): ANCHOR = XX2
    Case 2, 4
      percentage! = (X - XX2) / (XX1 - XX2): ANCHOR = XX1
    Case Else
      BezierStore
      Exit Sub
End Select

Boat1 = ((ANCHOR - Boat1) * percentage!) + Boat1
Boat2 = ((ANCHOR - Boat2) * percentage!) + Boat2
Boat3 = ((ANCHOR - Boat3) * percentage!) + Boat3
Boat4 = ((ANCHOR - Boat4) * percentage!) + Boat4

p1.X = Boat1: P1X(B) = Boat1: p2.X = Boat2: P2X(B) = Boat2
h1.X = Boat3: h1X(B) = Boat3: h2.X = Boat4: h2X(B) = Boat4
BezierStore

End Sub

Private Sub BezUpdate()

OFBN = 45 + NodeZoom

p1.X = P1X(B): p1.Y = P1Y(B)
p2.X = P2X(B): p2.Y = p2Y(B)
h1.X = h1X(B): h1.Y = h1Y(B)
h2.X = h2X(B): h2.Y = h2Y(B)
BezHandle1.Left = h1.X - OFBN: BezHandle1.Top = h1.Y - OFBN
BezHandle2.Left = h2.X - OFBN: BezHandle2.Top = h2.Y - OFBN
BezNode1.Left = p1.X - OFBN: BezNode1.Top = p1.Y - OFBN
BezNode2.Left = p2.X - OFBN: BezNode2.Top = p2.Y - OFBN
BezHandle1.ScaleLeft = BezHandle1.Left
BezHandle1.ScaleTop = BezHandle1.Top
BezHandle2.ScaleLeft = BezHandle2.Left
BezHandle2.ScaleTop = BezHandle2.Top
BezNode1.ScaleLeft = BezNode1.Left
BezNode1.ScaleTop = BezNode1.Top
BezNode2.ScaleLeft = BezNode2.Left
BezNode2.ScaleTop = BezNode2.Top
Shape1.Left = h1.X - OFBN: Shape1.Top = h1.Y - OFBN
Shape2.Left = h2.X - OFBN: Shape2.Top = h2.Y - OFBN
ReScale ''' If ZoomTool Used

ArtWork.DrawMode = 7: ArtWork.DrawStyle = 2
ArtWork.Line (p1.X, p1.Y)-(h1.X, h1.Y), YELLOW
ArtWork.Line (p2.X, p2.Y)-(h2.X, h2.Y), YELLOW
ArtWork.DrawStyle = 0

If ShowDotted = False Then VState = False Else VState = True
BezHandle1.Visible = VState: BezHandle2.Visible = VState
If Not CurrentTool = 1 Then VState = False Else VState = True
BezNode1.Visible = VState: BezNode2.Visible = VState

If CurrentTool = 1 Then N_H_Clear Else Nodes

End Sub

Private Sub BoundingBox_Down(X As Single, Y As Single)

For R = -30 To 30 Step 15: For RR = -30 To 30 Step 15
If ArtWork.Point(X - R, Y - RR) <> ArtWork.BackColor Then
   FlagMove = True
   XCenter = X - ((W(B) \ 2) + (E(B)) \ 2)
   YCenter = Y - ((N(B) \ 2) + (S(B)) \ 2)
   
   OFFSETx = XCenter: OFFSETy = YCenter: Exit For
End If
'ArtWork.PSet (X - R, Y - RR) 'use To view point area
If FlagMove = True Then Exit For
Next RR: Next R

    X1 = X: Y1 = Y: X2 = X1: Y2 = Y1
    Radius = 0
    ArtWork.DrawMode = INVERSE
    XX1 = W(B)
    YY1 = N(B)
    XX2 = E(B)
    YY2 = S(B)

End Sub

Private Sub BoundingBox_Move(Shift As Integer, X As Single, Y As Single)

If CurrentTool = 0 And FlagMove = True And MoveMouse() = True Then MoveObject Button, Shift, X, Y: Exit Sub
If Node(0).Visible = True Then Exit Sub

If MoveMouse() = True Then
Select Case CurrentTool
    Case PointerTool To BezierTool
      ArtWork.DrawMode = Xor_Pen
      ArtWork.DrawStyle = DOT
      ArtWork.Line (X1, Y1)-(X2, Y2), YELLOW, B
    Case LineTool
      ArtWork.Line (X1, Y1)-(X2, Y2)
    Case BoxTool
      ArtWork.Line (X1, Y1)-(X2, Y2), , B
    Case CircleTool
      LastCircleGood = PerfectCircle(X1, Y1, X2, Y2)
      ArtWork.Circle (XCenter, YCenter), Radius, , , , Aspect
    End Select
End If

End Sub

Private Sub BoundingBox_Up(X1&, Y1&, X2&, Y2&)

If MoveMouse() = True Then
Total = Total + 1: B = Total
StatusLabel.Caption = Format$(B)

LineWidth(B) = ArtWork.DrawWidth
W(B) = X1: E(B) = X2: N(B) = Y1: S(B) = Y2

Select Case CurrentTool
    Case PointerTool
        Exit Sub
    Case BezierTool
        Exit Sub
    Case ZoomTool
        Exit Sub
    Case LineTool
        FlagLine(B) = B
        Nodes
    Case BoxTool
        FlagBox(B) = B
        FillValue(B) = 1
        Nodes
    Case CircleTool
        FlagCircle(B) = B
        FillValue(B) = 1
        Nodes
    Case TextTool
        ''''''''''''
    End Select
End If

End Sub

Private Sub Color_INI()

ColorString = Space$(144)

R = GetPrivateProfileString(KEY_NAME, KEY_ENTRY + "1", BAS_PALETTE1, ColorString, Len(ColorString), INI_FILENAME)
  PaletteBarColors ColorString, 0

R = GetPrivateProfileString(KEY_NAME, KEY_ENTRY + "2", BAS_PALETTE2, ColorString, Len(ColorString), INI_FILENAME)
  PaletteBarColors ColorString, 16

R = GetPrivateProfileString(KEY_NAME, KEY_ENTRY + "3", BAS_PALETTE3, ColorString, Len(ColorString), INI_FILENAME)
  PaletteBarColors ColorString, 32

End Sub

Private Sub ColorBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (X >= 0) And (X <= 48) Then Update_Mouse_Colors Button, X, Y

End Sub

Private Sub ColorBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RulerGuides Index, Button, Shift, X, Y

End Sub

Private Sub ColorBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Node(0).Visible = False Then Exit Sub

Select Case Button
    Case RIGHT_BUTTON
      OutLineColor(B) = PenColor

    Case LEFT_BUTTON
      PaintColor(B) = BrushColor: FillValue(B) = 0
End Select

UpdateObjects

End Sub

Private Sub ColorBar_Paint()

    Display_Color_Palette ColorBar
    
End Sub

Private Sub ColorScroll_Change()

    ColorBar.Move ColorScroll.value

End Sub

Private Sub Find_Object(X As Single, Y As Single)

R = TGL
Do While R > 0 '''Find GuideLines

If X < (GuideX(R) + 30) And X > (GuideX(R) - 30) Then
    ArtWork.DrawMode = Xor_Pen: ArtWork.DrawStyle = DOT
    ArtWork.AutoRedraw = True
    ArtWork.Line (-5000, -5000)-(GuideX(R), 24000), YELLOW, B
    ArtWork.AutoRedraw = False: X = -5000: SideRuler = True: V = R
    Exit Sub
End If


If Y < (GuideY(R) + 30) And Y > (GuideY(R) - 30) Then
     ArtWork.DrawMode = Xor_Pen: ArtWork.DrawStyle = DOT
     ArtWork.AutoRedraw = True
     ArtWork.Line (-5000, -5000)-(24000, GuideY(R)), YELLOW, B
     ArtWork.AutoRedraw = False: Y = -5000: TopRuler = True: V = R
     Exit Sub
End If
R = R - 1
Loop
''''''''''Find Bounding Box
i = Total
Do While i > 0
If S(i) >= N(i) And E(i) >= W(i) Then West = W(i): East = E(i): North = N(i): South = S(i) 'NORMAL
If N(i) >= S(i) And W(i) >= E(i) Then West = E(i): East = W(i): North = S(i): South = N(i) 'DIAGONAL
If N(i) >= S(i) And W(i) <= E(i) Then West = W(i): East = E(i): North = S(i): South = N(i) 'top/bottom
If S(i) >= N(i) And W(i) >= E(i) Then West = E(i): East = W(i): North = N(i): South = S(i) 'side to side
If X >= West And X <= East And Y >= North And Y <= South Then: Exit Do
  i = i - 1
  Loop

If i = 0 Then PtInRectBasic = False: Exit Sub
  B = i: PtInRectBasic = True
  StatusLabel.Caption = Format$(B)

  Artisan.StatusColor.BackColor = OutLineColor(B)
  Artisan.StatusPaint.BackColor = PaintColor(B)
  Nodes

End Sub

Private Sub FindAllObjects(X1&, Y1&, X2&, Y2&, X As Single, Y As Single)

CopyX1 = X1: CopyX2 = X2: CopyY1 = Y1: CopyY2 = Y2

If X2 > X1 And Y2 > Y1 Then X1 = CopyX1: X2 = CopyX2: Y1 = CopyY1: Y2 = CopyY2 'NW >SE
If X2 > X1 And Y1 > Y2 Then X1 = CopyX1: X2 = CopyX2: Y1 = CopyY2: Y2 = CopyY1 'SW >NE
If X1 > X2 And Y1 > Y2 Then X1 = CopyX2: X2 = CopyX1: Y1 = CopyY2: Y2 = CopyY1 'SE >NW
If X1 > X2 And Y2 > Y1 Then X1 = CopyX2: X2 = CopyX1: Y1 = CopyY1: Y2 = CopyY2 'NE >SW

i = Total
Do While i > 0

If E(i) >= W(i) And S(i) >= N(i) Then West = W(i): East = E(i): North = N(i): South = S(i) 'NW >SE
If E(i) >= W(i) And N(i) >= S(i) Then West = W(i): East = E(i): North = S(i): South = N(i) 'SW >NE
If W(i) >= E(i) And N(i) >= S(i) Then West = E(i): East = W(i): North = S(i): South = N(i) 'SE >NW
If W(i) >= E(i) And S(i) >= N(i) Then West = E(i): East = W(i): North = N(i): South = S(i) 'NE >SW

If X1 <= West And Y1 <= North And X2 >= East And Y2 >= South Then: B = i: Handles: Nodes
i = i - 1
Loop

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If CurrentTool <> 8 Then
If KeyCode = KEY_DELETE Then Menu_EditSelection_click 6
If KeyCode = KEY_SPACE Then ToolButtons_Paint: CurrentTool = 0: ToolButtons_Paint
End If

End Sub

Private Sub Form_Load()
    
PixelX = Screen.TwipsPerPixelX: PixelY = Screen.TwipsPerPixelY
WIDE = Screen.Width \ PixelX: HIGH = Screen.Height \ PixelY
    
    NewGuideLine = False
    ArtWork.FontSize = 12: Txt.FontSize = 12
    ArtWork.FontName = Txt.FontName
    WindowState = 2
    TextPoint = 1
    PerSH = 1: PerSW = 1
    StatusColor.BackColor = BLACK
    LineWidth(B) = 1
    PenColor = BLACK
    MoveFlag = False
    LastMenu = 3
    Convert_Curve = False
    Convert_Line = False
    p = 1: B = 0: Total = 0
    ArtWork.Enabled = False
    ArtWork.BackColor = WHITE
    ArtWork.FillColor = WHITE
    ArtWork.FillStyle = TRANSPARENT
    VScroll.value = 0: HScroll.value = 0
    Color_INI
    Load TextForm
    KeyPreview = True
    
End Sub

Private Sub Form_Resize()

ToolButtons.ZOrder: StatusArea.ZOrder: ColorMask.ZOrder

'StatusArea picture box''''''
StatusArea.Visible = False
StatusArea.Move 0, 0, WIDE, 70
ColorMask.Move 0, 50, WIDE, 20
ColorBar.Move 0, 0, WIDE, 20
StatusColor.Move WIDE - 36, 20, 28, 28
StatusPaint.Move 2, 2, 24, 24
XYLabel.Move 36, 20, WIDE \ 5, 26
StatusLabel.Move (WIDE \ 2) - (StatusLabel.Width \ 2), 20, WIDE \ 3, 26
ColorScroll.Move 0, 20, 34, 17
No_Fill.Move 0, 37, 34, 13
TestButton.Move 0, 0, 56, 17
StatusArea.Visible = True
''''''''''''''''''''''''''''

ZoomBar.Move 64, ToolButtons.Top + 70, 128, 34
LineButtons.Move 64, ToolButtons.Top + 192, 256, 34
PaintButtons.Move 64, ToolButtons.Top + 224, 256, 34

PictureMask.Move -200, -250, WIDE + 400, HIGH + 400
ArtWork.BorderStyle = 0
ArtWork.Move 0, 0, PictureMask.Width, PictureMask.Height

RulerCross.Move 40, 0, 22, 22
Ruler(0).Move 40, PictureMask.Top, 22, PictureMask.Height 'Side Ruler
Ruler(1).Move PictureMask.Left, 0, PictureMask.Width, 22 'Top ruler

Txt.Move 0, 0
ToolButtons.Move 0, 0, 40, HIGH
'''''''''''

VScroll.Move WIDE - 17, 22, 17, StatusArea.Top - 22
HScroll.Move 56, 0, WIDE - 73, 17

ScaleBox.Move 3600, 3750, 330, 330

OldSW = ArtWork.ScaleWidth: OldSH = ArtWork.ScaleHeight
OldRW = 300: OldRH = 300
Ruler(0).Scale: Ruler(1).Scale

UpdateObjects
CurrentTool = 0

End Sub

Private Sub GuideLines(Index As Integer, Button As Integer, X As Single, Y As Single)

ArtWork.DrawStyle = 2: ArtWork.DrawMode = 7

If SideRuler = True Then  'Side Ruler
   ArtWork.Line (-5000, -5000)-(X2, Y2), YELLOW, B
   X2 = X
   Y2 = 40000
   ArtWork.Line (-5000, -5000)-(X2, Y2), YELLOW, B
End If

If TopRuler = True Then  'Top Ruler
    ArtWork.Line (-5000, -5000)-(X2, Y2), YELLOW, B
    X2 = 40000
    Y2 = Y
    ArtWork.Line (-5000, -5000)-(X2, Y2), YELLOW, B
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Handle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If B = FlagText(B) Then Exit Sub

ArtWork.DrawMode = Xor_Pen
ArtWork.DrawWidth = 1
    N_H_Clear
    XX1 = W(B)
    YY1 = N(B)
    XX2 = E(B)
    YY2 = S(B)
ArtWork.Line (W(B), N(B))-(E(B), S(B)), YELLOW, B

End Sub

Private Sub Handle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If B = FlagText(B) Then Exit Sub
    Xstr = Format$(X \ PixelX, "000")
    Ystr = Format$(Y \ PixelY, "000")
    XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"
RulerGuides Index, Button, Shift, X, Y

If Button Then
    ArtWork.Line (W(B), N(B))-(E(B), S(B)), YELLOW, B
    ArtWork.DrawStyle = DOT

Select Case Index
  Case 0
    Handle(0).MousePointer = 8
    W(B) = X
    N(B) = Y
  Case 1
    Handle(1).MousePointer = 7
    N(B) = Y
  Case 2
    Handle(2).MousePointer = 6
    E(B) = X
    N(B) = Y
  Case 3
    Handle(3).MousePointer = 9
    E(B) = X
  Case 4
    Handle(4).MousePointer = 8
    E(B) = X
    S(B) = Y
  Case 5
    Handle(5).MousePointer = 7
    S(B) = Y
  Case 6
    Handle(6).MousePointer = 6
    W(B) = X
    S(B) = Y
   Case 7
    Handle(7).MousePointer = 9
    W(B) = X
End Select
ArtWork.Line (W(B), N(B))-(E(B), S(B)), YELLOW, B
End If

BoxWidth = E(B) - W(B): BoxHeight = S(B) - N(B)
If Button Then StatusLabel.Caption = "Width  " + Format$(BoxWidth) + " " + "Height  " + Format$(BoxHeight)

End Sub

Private Sub Handle_Mouseup(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If B = FlagText(B) Then Exit Sub
If FlagBezier(B) = B Then BezStretch Index, Button, Shift, X, Y

UpdateObjects
    
    For R = 0 To 7
    Handle(R).MousePointer = 2
    Next R

End Sub

Private Sub Handles()
BezNode1.Visible = False: BezNode2.Visible = False
Shape1.Visible = False: Shape2.Visible = False

If N(B) > S(B) And W(B) < E(B) Or S(B) > N(B) And W(B) > E(B) Then HandlesReversed: Exit Sub


If S(B) > N(B) And E(B) > W(B) Then AA = -120: BB = 88 'NORMAL
If N(B) > S(B) And W(B) > E(B) Then AA = 88: BB = -120 'DIAGONAL

OSH = 44 + NodeZoom 'OFFSET Handles
WC = W(B) - OSH: NC = N(B) - OSH: EC = E(B) - OSH: SC = S(B) - OSH

MH = (WC \ 2 + EC \ 2)
MV = (NC \ 2 + SC \ 2)

Handle(0).Left = WC + AA
Handle(0).Top = NC + AA
Handle(1).Left = MH
Handle(1).Top = NC + AA
Handle(2).Left = EC + BB
Handle(2).Top = NC + AA
Handle(3).Left = EC + BB
Handle(3).Top = MV
Handle(4).Left = EC + BB
Handle(4).Top = SC + BB
Handle(5).Left = MH
Handle(5).Top = SC + BB
Handle(6).Left = WC + AA
Handle(6).Top = SC + BB
Handle(7).Left = WC + AA
Handle(7).Top = MV

For R = 0 To 7
Handle(R).ScaleLeft = Handle(R).Left
Handle(R).ScaleTop = Handle(R).Top
Handle(R).ScaleWidth = Handle(R).Width + Zoom
Handle(R).ScaleHeight = Handle(R).Height + Zoom
Handle(R).Visible = True
Next R

End Sub

Private Sub HandlesReversed()

If N(B) > S(B) And W(B) < E(B) Then AA = 88: BB = -120 'top/bottom
If S(B) > N(B) And W(B) > E(B) Then AA = -120: BB = 88 'side to side

OSH = 44 + NodeZoom 'OFFSET Handles
WC = W(B) - OSH: NC = N(B) - OSH: EC = E(B) - OSH: SC = S(B) - OSH

MH = (WC \ 2 + EC \ 2)
MV = (NC \ 2 + SC \ 2)

Handle(0).Left = WC + BB
Handle(0).Top = NC + AA
Handle(1).Left = MH
Handle(1).Top = NC + AA
Handle(2).Left = EC + AA
Handle(2).Top = NC + AA
Handle(3).Left = EC + AA
Handle(3).Top = MV
Handle(4).Left = EC + AA
Handle(4).Top = SC + BB
Handle(5).Left = MH
Handle(5).Top = SC + BB
Handle(6).Left = WC + BB
Handle(6).Top = SC + BB
Handle(7).Left = WC + BB
Handle(7).Top = MV

For R = 0 To 7
Handle(R).ScaleLeft = Handle(R).Left
Handle(R).ScaleTop = Handle(R).Top
Handle(R).ScaleWidth = Handle(R).Width + Zoom
Handle(R).ScaleHeight = Handle(R).Height + Zoom
Handle(R).Visible = True
Next R

End Sub

Private Sub HScroll_Change()

ArtWork.Enabled = False
Dim WM_MOUSEMOVE%
WM_MOUSEMOVE% = &H200
API = SendMessageByNum(ArtWork.hwnd, WM_MOUSEMOVE%, 0, 0)

 ArtWork.Move -HScroll.value, ArtWork.Top
 Ruler(1).Move ArtWork.Left - 200, Ruler(1).Top
 BSM = (HScroll.value * PixelX) * PerSW

End Sub

Private Sub LineButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Node(0).Visible = False Then LineButtons.Visible = False: Exit Sub
CurrentTool = 6: ToolButtons_Paint: CurrentTool = LastToolUsed: ToolButtons_Paint
CurrentLine = X \ 32: LineButtons.Visible = False

Select Case CurrentLine
 Case LineColor
  CMDialog1.color = OutLineColor(B)
  CMDialog1.Flags = &H2& Or &H1&
  CMDialog1.Action = 3
  OutLineColor(B) = CMDialog1.color
 Case NoLine
    LineValue(B) = B
 Case HairLine
    ArtWork.DrawWidth = 1
 Case OnePt
    ArtWork.DrawWidth = 2
 Case TwoPt
    ArtWork.DrawWidth = 3
 Case FourPt
    ArtWork.DrawWidth = 5
 Case EightPt
    ArtWork.DrawWidth = 9
 Case TwelvePt
    ArtWork.DrawWidth = 13
End Select
LineWidth(B) = ArtWork.DrawWidth

If B = 0 Then Exit Sub
If Not CurrentLine = NoLine Then LineValue(B) = 0
UpdateObjects

End Sub

Private Sub LineButtons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ArtWork.DrawWidth = 1

End Sub

Private Sub Menu_Arrange_Click()

    LastMenu = 2

End Sub

Private Sub Menu_ArrangeSelection_Click(Index As Integer)

Select Case Index
    Case 0
     To_Front_Back Index
    Case 1
     To_Front_Back Index
End Select

End Sub

Private Sub Menu_Color_Click()

    LastMenu = 3

End Sub

Private Sub Menu_ColorSelection_Click(Index As Integer)

Select Case Index
    Case 2 'Show rulers
     Menu_ColorSelection(2).Checked = Not Menu_ColorSelection(2).Checked
     Y1Ruler = 0: X2Ruler = 0
     RulerCross.Visible = Not RulerCross.Visible
     Ruler(0).Visible = Not Ruler(0).Visible
     Ruler(1).Visible = Not Ruler(1).Visible
     
End Select

End Sub

Private Sub Menu_Edit_Click()

    LastMenu = 1

End Sub

Private Sub Menu_EditSelection_click(Index As Integer)
    
Select Case Index
    Case 0 'UNDO
    '''''''''''''''
    Case 1 'REDO
    '''''''''''''''
    Case 3 'CUT
    '''''''''''''''
    Case 4 'COPY
    '''''''''''''''
    Case 5 'PASTE
    '''''''''''''''
    Case 6 'DELETE
If B = 0 Then Exit Sub
If Node(0).Visible = True Then
N_H_Clear
FillValue(B) = 0: LineValue(B) = 0: OutLineColor(B) = 0
PaintColor(B) = 0: LineWidth(B) = 1: FlagAlign(B) = 0
OFFSET(B) = 0: TextSize(B) = 0: TextFont(B) = ""
TextBold(B) = 0: TextItalic(B) = 0: CarriageCount(B) = 0
Storedtext(B) = ""
FlagBezier(B) = 0: FlagLine(B) = 0: FlagBox(B) = 0
FlagCircle(B) = 0: FlagText(B) = 0
P1X(B) = 0: P1Y(B) = 0: P2X(B) = 0: p2Y(B) = 0
h1X(B) = 0: h1Y(B) = 0: h2X(B) = 0: h2Y(B) = 0

DeleteObject = True
XX1 = W(B): YY1 = N(B): XX2 = E(B): YY2 = S(B)
W(B) = 0: N(B) = 0: E(B) = 0: S(B) = 0
UpdateObjects
DeleteObject = False
End If
    
    Case 7  'SELECT ALL
    '''''''''''''''
    Case 9 'COPY DESIGN FROM
    '''''''''''''''
    Case 10 'EDIT TEXT
    '''''''''''''''
    Case 12 'ROTATE SKEW
    '''''''''''''''
    Case 13 'MIRROR
    '''''''''''''''
End Select
   
End Sub

Private Sub Menu_File_Click()

  LastMenu = 0
  N_H_Clear
  If Txt.Caption > "" Then ReDrawText

  If BezHandle1.Visible = True Then
  ShowDotted = True: BezUpdate
  BezHandle1.Visible = False: BezHandle2.Visible = False
  BezNode1.Visible = False: BezNode2.Visible = False
  Shape1.Visible = False: Shape2.Visible = False
  End If

End Sub

Private Sub Menu_FileSelection_Click(Index As Integer)

Select Case Index
  Case 0 'New
Erase FillValue, LineValue, OutLineColor, PaintColor, LineWidth
Erase FlagAlign, OFFSET, TextSize, TextFont, TextBold, TextItalic
Erase CarriageCount, Storedtext, W, E, N, S
Erase FlagBezier, FlagLine, FlagBox, FlagCircle, FlagText
Erase P1X, P1Y, P2X, p2Y, h1X, h1Y, h2X, h2Y
Erase GuideX, GuideY

p = 0: Total = 0: V = 0: B = 0: TGL = 0
     ArtWork.DrawWidth = 1
     ArtWork.DrawMode = 6: ArtWork.DrawStyle = 0
     ArtWork.AutoRedraw = True: ArtWork.Cls
     UpdateObjects
  Case 1 'OPEN
    mnuFileOpen
  Case 2 To 3
    mnuFileSave
  Case 4 'Print
    WinPrint
  Case 5 'EXIT
    Unload Me
     End
End Select

End Sub

Private Sub Menu_Palette_Click(Index As Integer)

Select Case Index
    Case 0      'Edit Color Palette
      ColorIndex = 0
      ColorPalette.Show
    Case 1      'Default Palette
      PaletteBarColors BAS_PALETTE1, 0
      PaletteBarColors BAS_PALETTE2, 16
      PaletteBarColors BAS_PALETTE3, 32
      Display_Color_Palette ColorBar
    Case 2
      Color_INI 'Custom Palette
      Display_Color_Palette ColorBar
    Case 3      'Save As Custom
      SaveColor_INI
End Select

End Sub

Private Sub mnuFileOpen()

On Error GoTo ErrHandler
CMDialog1.Flags = &H800&
CMDialog1.Filter = "All Files(*.*)|*.*|Art Files(*.art)|*.art"
CMDialog1.DefaultExt = "art"
CMDialog1.FilterIndex = 1
CMDialog1.Action = 1

Dim B1 As Variant
Open CMDialog1.FileName For Input As #1
Line Input #1, B1: Total = B1

For B = 1 To Total

Line Input #1, B1: FillValue(B) = B1
Line Input #1, B1: LineValue(B) = B1
Line Input #1, B1: OutLineColor(B) = B1
Line Input #1, B1: PaintColor(B) = B1
Line Input #1, B1: LineWidth(B) = B1
Line Input #1, B1: FlagAlign(B) = B1
Line Input #1, B1: OFFSET(B) = B1
Line Input #1, B1: TextSize(B) = B1
Line Input #1, B1: TextFont(B) = B1
Line Input #1, B1: TextBold(B) = B1
Line Input #1, B1: TextItalic(B) = B1
Line Input #1, B1: CarriageCount(B) = B1

If CarriageCount(B) = 0 Then
Line Input #1, B1: Storedtext(B) = B1
Else
For R = 0 To CarriageCount(B)
Line Input #1, B1: Alltext$ = Alltext$ + B1 + Chr(13): Storedtext(B) = Alltext$
Next R
Storedtext(B) = Left$(Alltext$, Len(Alltext$) - 1)
End If
Alltext$ = ""

Line Input #1, B1: W(B) = B1
Line Input #1, B1: E(B) = B1
Line Input #1, B1: N(B) = B1
Line Input #1, B1: S(B) = B1
Line Input #1, B1: FlagBezier(B) = B1
Line Input #1, B1: FlagLine(B) = B1
Line Input #1, B1: FlagBox(B) = B1
Line Input #1, B1: FlagCircle(B) = B1
Line Input #1, B1: FlagText(B) = B1
Line Input #1, B1: P1X(B) = B1
Line Input #1, B1: P1Y(B) = B1
Line Input #1, B1: P2X(B) = B1
Line Input #1, B1: p2Y(B) = B1
Line Input #1, B1: h1X(B) = B1
Line Input #1, B1: h1Y(B) = B1
Line Input #1, B1: h2X(B) = B1
Line Input #1, B1: h2Y(B) = B1

Next B

Total = B - 1: B = 0
UpdateObjects

Exit Sub

ErrHandler:
Close #1: Exit Sub

End Sub

Private Sub mnuFileSave()

On Error GoTo BadFileName:
CMDialog1.Flags = &H1802&
CMDialog1.Filter = "All Files(*.*)|*.*|Art Files(*.art)|*.art"
CMDialog1.DefaultExt = "art"
CMDialog1.FilterIndex = 1
CMDialog1.Action = 2

Open CMDialog1.FileName For Output As #1
0:  Print #1, Total

For B = 1 To Total
1:  Print #1, FillValue(B)
2:  Print #1, LineValue(B)
3:  Print #1, OutLineColor(B)
4:  Print #1, PaintColor(B)
5:  Print #1, LineWidth(B)
6:  Print #1, FlagAlign(B)
7:  Print #1, OFFSET(B)
8:  Print #1, TextSize(B)
9:  Print #1, TextFont(B)
10:  Print #1, TextBold(B)
11:  Print #1, TextItalic(B)
12:  Print #1, CarriageCount(B)
13:  Print #1, Storedtext(B)
14:  Print #1, W(B)
15:  Print #1, E(B)
16:  Print #1, N(B)
17:  Print #1, S(B)
18:  Print #1, FlagBezier(B)
19:  Print #1, FlagLine(B)
20:  Print #1, FlagBox(B)
21:  Print #1, FlagCircle(B)
22:  Print #1, FlagText(B)
23:  Print #1, P1X(B)
24:  Print #1, P1Y(B)
25:  Print #1, P2X(B)
26:  Print #1, p2Y(B)
27:  Print #1, h1X(B)
28:  Print #1, h1Y(B)
29:  Print #1, h2X(B)
30:  Print #1, h2Y(B)

Next B

BadFileName:
Close #1
Exit Sub

End Sub

Private Function MoveMouse()

If X1 > X2 + 45 Or X1 < X2 - 45 Or Y1 > Y2 + 45 Or Y1 < Y2 - 45 Then
  MoveMouse = True
Else
  MoveMouse = False
End If

End Function

Private Sub MoveObject(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If B = 0 Or PtInRectBasic = False Then Exit Sub

 N_H_Clear
 ArtWork.DrawMode = 7
 ArtWork.DrawWidth = 1
 'ArtWork.DrawStyle = Dot
 ArtWork.MousePointer = 5
 
 XCenter = X - ((W(B) \ 2) + (E(B)) \ 2)
 YCenter = Y - ((N(B) \ 2) + (S(B)) \ 2)
 ArtWork.Line (W(B), N(B))-(E(B), S(B)), YELLOW, B
 W(B) = (W(B) - OFFSETx) + XCenter
 N(B) = (N(B) - OFFSETy) + YCenter
 E(B) = (E(B) - OFFSETx) + XCenter
 S(B) = (S(B) - OFFSETy) + YCenter
 
 If FlagBezier(B) = B Then
    P1X(B) = (P1X(B) - OFFSETx) + XCenter
    P1Y(B) = (P1Y(B) - OFFSETy) + YCenter
    P2X(B) = (P2X(B) - OFFSETx) + XCenter
    p2Y(B) = (p2Y(B) - OFFSETy) + YCenter
    h1X(B) = (h1X(B) - OFFSETx) + XCenter
    h1Y(B) = (h1Y(B) - OFFSETy) + YCenter
    h2X(B) = (h2X(B) - OFFSETx) + XCenter
    h2Y(B) = (h2Y(B) - OFFSETy) + YCenter
 End If

 If FlagText(B) = B Then OFFSET(B) = (OFFSET(B) - OFFSETx) + XCenter

 ArtWork.Line (W(B), N(B))-(E(B), S(B)), YELLOW, B

End Sub

Private Sub N_H_Clear()
    
    If B > 0 Then
      For R = 0 To 7
        Handle(R).Visible = False
      Next R
    End If
    
    Node(0).Visible = False
    Node(1).Visible = False
    Node(2).Visible = False
    Node(3).Visible = False
    
End Sub

Private Sub No_Fill_Click()

If B = FlagLine(B) Or Node(0).Visible = False Then Exit Sub
  
  FillValue(B) = 1
  UpdateObjects

End Sub

Private Sub Node_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not FlagLine(B) = B And Not FlagBezier(B) = B Then Exit Sub
OSN = 45 + NodeZoom
Node(0).Visible = False: Node(2).Visible = False
ArtWork.DrawMode = INVERSE

Select Case CurrentTool
    Case BezierTool
      ArtWork.DrawMode = 13
      ArtWork.AutoRedraw = True
      ArtWork.Line (W(B), N(B))-(E(B), S(B)), WHITE
      ArtWork.AutoRedraw = False
      ArtWork.DrawMode = INVERSE
      ArtWork.Line (W(B), N(B))-(E(B), S(B)), WHITE

      If Index = 2 Then
      X1 = W(B): Y1 = N(B)
      X2 = E(B): Y2 = S(B)
      W(B) = E(B): N(B) = S(B)

      ElseIf Index = 0 Then
      X1 = E(B): Y1 = S(B)
      X2 = W(B): Y2 = N(B)
      End If
      XX2 = X2: YY2 = Y2: XX1 = X1: YY1 = Y1
    Case LineTool
      If Index = 0 Then X1 = Node(0).Left + OSN: Y1 = Node(0).Top + OSN
      If Index = 2 Then X1 = Node(2).Left + OSN: Y1 = Node(2).Top + OSN
      X2 = X1: Y2 = Y1
End Select

End Sub

Private Sub Node_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Xstr = Format$(X \ PixelX, "000")
Ystr = Format$(Y \ PixelY, "000")
XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"

If Not FlagLine(B) = B And Not FlagBezier(B) = B Then Exit Sub

If Button Then
Select Case CurrentTool
    Case BezierTool
      If FlagLine(B) = B Then
      ArtWork.Line (X1, Y1)-(X2, Y2)
      X1 = X: Y1 = Y
      ArtWork.Line (X1, Y1)-(X2, Y2)
      End If
    Case LineTool
      ArtWork.Line (X1, Y1)-(X2, Y2)
      X2 = X: Y2 = Y
      ArtWork.Line (X1, Y1)-(X2, Y2)
End Select
End If

End Sub

Private Sub Node_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not FlagLine(B) = B And Not FlagBezier(B) = B Then Exit Sub

Select Case CurrentTool
    Case BezierTool
      W(B) = X1: N(B) = Y1: E(B) = X2: S(B) = Y2
      UpdateObjects
      Nodes
    Case LineTool
      ArtWork.DrawMode = COPY_PEN
      BoundingBox_Up X1, Y1, X2, Y2
      If B = Total Then Draw = True: UpdateObjects
    End Select
    
End Sub

Private Sub Nodes()

OSN = 45 + NodeZoom 'OFFSET Nodes

Select Case B
  Case FlagLine(B)
    Node(0).Left = E(B) - OSN
    Node(0).Top = S(B) - OSN
    Node(2).Left = W(B) - OSN
    Node(2).Top = N(B) - OSN
    Node(0).Visible = True
    Node(1).Visible = False
    Node(2).Visible = True
    Node(3).Visible = False
  
  Case FlagBezier(B)
    Node(0).Left = P1X(B) - OSN
    Node(0).Top = P1Y(B) - OSN
    Node(1).Left = P2X(B) - OSN
    Node(1).Top = p2Y(B) - OSN
    Node(0).Visible = True
    Node(1).Visible = True
    
  Case FlagCircle(B)
    p = (W(B) + E(B)) \ 2
    Node(0).Left = p - OSN
    Node(0).Top = N(B) - OSN
Node(0).Visible = True
Node(1).Visible = False
Node(2).Visible = False
Node(3).Visible = False

  Case FlagBox(B)
    Node(0).Left = E(B) - OSN
    Node(0).Top = S(B) - OSN
    Node(1).Left = W(B) - OSN
    Node(1).Top = S(B) - OSN
    Node(2).Left = W(B) - OSN
    Node(2).Top = N(B) - OSN
    Node(3).Left = E(B) - OSN
    Node(3).Top = N(B) - OSN
For R = 0 To 3
 Node(R).Visible = True
Next R
  Case FlagText(B)
    Node(0).Left = W(B) + Zoom
    Node(0).Top = S(B) + Zoom
    Node(0).Visible = True
End Select

ReScale

End Sub

Private Sub PaintButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Node(0).Visible = False Then PaintButtons.Visible = False: Exit Sub
CurrentTool = 7: ToolButtons_Paint: CurrentTool = LastToolUsed: ToolButtons_Paint
CurrentBrush = X \ 32: PaintButtons.Visible = False

Select Case CurrentBrush
 Case Palette
  CMDialog1.color = PaintColor(B)
  CMDialog1.Flags = &H2& Or &H1&
  CMDialog1.Action = 3
  PaintColor(B) = CMDialog1.color: FillValue(B) = 0
 Case NoFill
  No_Fill_Click
 Case WhiteBrush
  PaintColor(B) = &HFFFFFF: FillValue(B) = 0
 Case Brush10
  PaintColor(B) = &HE5E5E5: FillValue(B) = 0
 Case Brush20
  PaintColor(B) = &HCCCCCC: FillValue(B) = 0
 Case Brush50
  PaintColor(B) = &H7F7F7F: FillValue(B) = 0
 Case Brush70
  PaintColor(B) = &H4C4C4C: FillValue(B) = 0
 Case BlackBrush
  PaintColor(B) = &H0: FillValue(B) = 0
End Select
UpdateObjects

End Sub

Private Sub PaletteBarColors(ColorString As String, FirstColor)
    
 Start = 1
 For C = FirstColor To FirstColor + 15
   Colors(C) = Val(Mid$(ColorString, Start, 8))
   Start = Start + 9
 Next C

End Sub

Private Function PerfectCircle(X1&, Y1&, X2&, Y2&)

Xleg = Abs(X2 - X1)
    If Xleg <> 0 Then
        Yleg = Abs(Y2 - Y1)
        Aspect = Abs(Yleg / Xleg)
        If Xleg > Yleg Then Radius = Xleg \ 2 Else Radius = Yleg \ 2
        XCenter = X1 + (X2 - X1) \ 2
        YCenter = Y1 + (Y2 - Y1) \ 2
    End If

    PerfectCircle = Xleg

End Function

Private Sub ReDrawText()

If Txt.Caption > "" Then
Total = Total + 1: B = Total

CarriageCount(B) = crCounter: crCounter = 0

LineWidth(B) = ArtWork.DrawWidth
FlagText(B) = B: FillValue(B) = 1

TL = Txt.Left: TT = Txt.Top
TW = Txt.Left + Txt.Width
TH = (Txt.Top + Txt.Height)
Txt.Visible = False

TextSize(B) = Txt.FontSize ' Size of Txt.Label Type
Storedtext(B) = RTrim(Txt.Caption)
TextFont(B) = RTrim(ArtWork.FontName)
TextBold(B) = ArtWork.FontBold
TextItalic(B) = ArtWork.FontItalic

If Align = 0 Then FlagAlign(B) = TA_LEFT: OS = 0
If Align = 6 Then FlagAlign(B) = TA_CENTER: OS = (Txt.Width / 2)
If Align = 2 Then FlagAlign(B) = TA_RIGHT: OS = Txt.Width
''''''''''''''''''''''''''''''''''''
    ArtWork.AutoRedraw = True
    ArtWork.CurrentX = TL
    OFFSET(B) = ArtWork.CurrentX
    ArtWork.CurrentY = TT
    SaveScaleLeft = ArtWork.ScaleLeft
    ArtWork.ScaleLeft = -ArtWork.CurrentX + Scaling
    ArtWork.ForeColor = PaintColor(B)
    ArtWork.FontName = TextFont(B)
    ArtWork.FontSize = TextSize(B)
    TextSize(B) = Txt.FontSize / TextPoint
    ArtWork.FontBold = TextBold(B)
    ArtWork.FontItalic = TextItalic(B)
    wFlag = FlagAlign(B)
    TA = SetTextAlign(ArtWork.hDC, wFlag)
    ArtWork.Print Storedtext(B)
    ArtWork.AutoRedraw = False
    ArtWork.ScaleLeft = SaveScaleLeft

W(B) = TL - OS: E(B) = TW - OS: N(B) = TT: S(B) = TH
''''''''''''''''''''''''''
Txt.Caption = ""
End If

  Txt.Visible = False
  Txt.FontBold = False
  Txt.FontItalic = False
  Txt.Width = 15

  ArtWork.FontBold = False
  ArtWork.FontItalic = False
  Align = 0

End Sub

Private Sub ReScale()

 BezHandle1.ScaleWidth = BezHandle1.Width + Zoom
 BezHandle1.ScaleHeight = BezHandle1.Height + Zoom
 BezHandle2.ScaleWidth = BezHandle2.Width + Zoom
 BezHandle2.ScaleHeight = BezHandle2.Height + Zoom
 BezNode1.ScaleWidth = BezNode1.Width + Zoom
 BezNode1.ScaleHeight = BezNode1.Height + Zoom
 BezNode2.ScaleWidth = BezNode2.Width + Zoom
 BezNode2.ScaleHeight = BezNode2.Height + Zoom
 ''''''
For R = 0 To 3
Node(R).ScaleLeft = Node(R).Left
Node(R).ScaleTop = Node(R).Top
Node(R).ScaleWidth = Node(R).Width + Zoom
Node(R).ScaleHeight = Node(R).Height + Zoom
Next R

For R = 0 To 7
Handle(R).ScaleLeft = Handle(R).Left
Handle(R).ScaleTop = Handle(R).Top
Handle(R).ScaleWidth = Handle(R).Width + Zoom
Handle(R).ScaleHeight = Handle(R).Height + Zoom
Next R

Zoom = 0

End Sub

Private Sub Ruler_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If BezHandle1.Visible = True Then ShowDotted = False: BezUpdate

NewGuideLine = True
X2 = -5000: Y2 = -5000

If Index = 0 Then SideRuler = True: Exit Sub
If Index = 1 Then TopRuler = True: Exit Sub

End Sub

Private Sub Ruler_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Xstr = Format$(X \ PixelX, "000")
    Ystr = Format$(Y \ PixelY, "000")

XYLabel.Caption = "[" + Xstr + " , " + Ystr + "]"
    
    RulerGuides Index, Button, Shift, X, Y
    If Button Then GuideLines Index, Button, X, Y

End Sub

Private Sub Ruler_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    X2 = X: Y2 = Y
    If NewGuideLine = True Then TGL = TGL + 1: V = TGL
    If SideRuler = True Then GuideX(V) = X2
    If TopRuler = True Then GuideY(V) = Y2
    
    NewGuideLine = False
    
    ArtWork.DrawStyle = SOLID
    UpdateObjects

SideRuler = False: TopRuler = False: ShowDotted = False

End Sub

Private Sub Ruler_Paint(Index As Integer)
    
Ruler(0).Cls: Ruler(1).Cls
Ruler(0).DrawMode = 13: Ruler(1).DrawMode = 13
Ruler(0).DrawStyle = SOLID: Ruler(1).DrawStyle = SOLID
'''''''''''''''''''''''''''''''''
Ruler(0).ScaleLeft = 0: Ruler(0).ScaleWidth = 4: Num = -8
  For i = 300 To 15000 Step 576
    Ruler(0).Line (0, i)-(4, i)
    Ruler(0).Line (2, i - 288)-(4, i - 288)
    Ruler(0).Line (3, i - 144)-(4, i - 144)
    Ruler(0).Line (3, i - 432)-(4, i - 432)
    Ruler(0).PSet (-0.5, i), Ruler(0).BackColor
    Num = Num + 1: Ruler(0).Print Num
  Next i
''''''''''''''''''''''''''''''''''
Ruler(1).ScaleTop = 0: Ruler(1).ScaleHeight = 4: Num = -12
  For i = 200 To 18000 Step 576
    Ruler(1).Line (i, 0)-(i, 4)
    Ruler(1).Line (i - 288, 0)-(i - 288, 2)
    Ruler(1).Line (i - 144, 0)-(i - 144, 1)
    Ruler(1).Line (i - 432, 0)-(i - 432, 1)
    Ruler(1).PSet (i, 1), Ruler(1).BackColor
    Num = Num + 1: Ruler(1).Print Num
  Next i
''''''''''''''''''''''''''''''''''
    Ruler(0).ScaleTop = ArtWork.ScaleTop
    Ruler(0).ScaleHeight = ArtWork.ScaleHeight
    Ruler(0).ScaleLeft = ScaleBox.Left + BSM
    Ruler(0).ScaleWidth = OldRW * PerSW

    Ruler(1).ScaleLeft = ArtWork.ScaleLeft
    Ruler(1).ScaleWidth = ArtWork.ScaleWidth
    Ruler(1).ScaleTop = ScaleBox.Top + SSM
    Ruler(1).ScaleHeight = OldRH * PerSH
'''''''''''''''''''''''''''''''''''
Ruler(0).DrawMode = Xor_Pen: Ruler(1).DrawMode = Xor_Pen
Ruler(0).DrawStyle = DOT: Ruler(1).DrawStyle = DOT
XRuler = -5000: YRuler = -5000
ArtWork.Enabled = True

End Sub

Private Sub RulerCross_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RulerGuides Index, Button, Shift, X, Y

End Sub

Private Sub RulerGuides(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

R0SL = Ruler(0).ScaleLeft 'Side Ruler
R0L = Ruler(0).Left - 5000: R0T = Ruler(0).Top - 5000
R1ST = Ruler(1).ScaleTop 'Top Ruler
R1L = Ruler(1).ScaleLeft - 15: R1T = Ruler(1).ScaleTop - 15

Ruler(0).Line (R0L, R0T)-(X1Ruler, Y1Ruler), CYAN, B 'Vert
Ruler(1).Line (R1L, R1T)-(X2Ruler, Y2Ruler), CYAN, B 'Horz
    X1Ruler = R0SL + 500  'Side Ruler.Width
    Y1Ruler = Y
    X2Ruler = X
    Y2Ruler = R1ST + 500  'Top Ruler.Height
Ruler(0).Line (R0L, R0T)-(X1Ruler, Y1Ruler), CYAN, B  'Vert
Ruler(1).Line (R1L, R1T)-(X2Ruler, Y2Ruler), CYAN, B  'Horz

End Sub

Private Sub SaveColor_INI()
    
For i = 0 To 2
  ColorString = ""
  For C = 0 To 15
    ColorString = ColorString + Format$(Colors(i * 16 + C), "00000000 ")
  Next C
  X = WritePrivateProfileString(KEY_NAME, KEY_ENTRY + Format$(i + 1), ColorString, INI_FILENAME)
Next i

End Sub

Private Sub StatusArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RulerGuides Index, Button, Shift, X, Y

End Sub

Private Sub StatusArea_Paint()

StatusArea.Cls

Select Case True
  Case Zoom_In
    PageSize$ = "76%"
  Case Zoom_Out
    PageSize$ = "22%"
  Case Else
    PageSize$ = "44%"

End Select

StatusArea.CurrentX = (StatusArea.Width * 0.75)
StatusArea.CurrentY = (StatusArea.Height \ 3)
StatusArea.Print PageSize$

End Sub

Private Sub To_Front_Back(Index As Integer)

Dim PC As Long: Dim OLC As Long
Dim St As String: Dim TF As String

If Node(0).Visible = False Then Exit Sub
If Index = 0 And B = Total Then Exit Sub
If Index = 1 And B <= 1 Then Exit Sub

''''''store objects attributes
If FlagText(B) = B Then FT = 1
If FlagCircle(B) = B Then FC = 1
If FlagLine(B) = B Then FL = 1
If FlagBox(B) = B Then FB = 1
If FlagBezier(B) = B Then Bez = 1

OLC = OutLineColor(B): PC = PaintColor(B): LW = LineWidth(B)
Fill = FillValue(B): LV = LineValue(B): St = Storedtext(B)
TS = TextSize(B): TF = TextFont(B): TB = TextBold(B): TI = TextItalic(B)
FA = FlagAlign(B): OF = OFFSET(B): CR = CarriageCount(B)

W1 = W(B): N2 = N(B): E3 = E(B): S4 = S(B)
p1XX = P1X(B): p1YY = P1Y(B): p2XX = P2X(B): p2YY = p2Y(B)
h1XX = h1X(B): h1YY = h1Y(B): h2XX = h2X(B): h2YY = h2Y(B)
'''''''End of every possible attribute

If Index = 0 Then Last = Total - 1: First = 1 Else Last = 2: First = -1
For B = B To Last Step First: A = B + (First)

W(B) = W(A): N(B) = N(A): E(B) = E(A): S(B) = S(A)
P1X(B) = P1X(A): P1Y(B) = P1Y(A)
P2X(B) = P2X(A): p2Y(B) = p2Y(A)
h1X(B) = h1X(A): h1Y(B) = h1Y(A)
h2X(B) = h2X(A): h2Y(B) = h2Y(A)

OutLineColor(B) = OutLineColor(A)
PaintColor(B) = PaintColor(A)
LineWidth(B) = LineWidth(A)
FillValue(B) = FillValue(A)
LineValue(B) = LineValue(A)
Storedtext(B) = Storedtext(A)
TextSize(B) = TextSize(A)
TextFont(B) = TextFont(A)
TextBold(B) = TextBold(A)
TextItalic(B) = TextItalic(A)
FlagAlign(B) = FlagAlign(A)
OFFSET(B) = OFFSET(A)
CarriageCount(B) = CarriageCount(A)

If FlagText(A) = 0 Then FlagText(B) = 0 Else FlagText(B) = B
If FlagCircle(A) = 0 Then FlagCircle(B) = 0 Else FlagCircle(B) = B
If FlagLine(A) = 0 Then FlagLine(B) = 0 Else FlagLine(B) = B
If FlagBox(A) = 0 Then FlagBox(B) = 0 Else FlagBox(B) = B
If FlagBezier(A) = 0 Then FlagBezier(B) = 0 Else FlagBezier(B) = B
Next B
 
''''Reset objects attributes
If FT = 1 Then FlagText(B) = B Else FlagText(B) = 0
If FB = 1 Then FlagBox(B) = B Else FlagBox(B) = 0
If FC = 1 Then FlagCircle(B) = B Else FlagCircle(B) = 0
If FL = 1 Then FlagLine(B) = B Else FlagLine(B) = 0
If Bez = 1 Then FlagBezier(B) = B Else FlagBezier(B) = 0
OutLineColor(B) = OLC: PaintColor(B) = PC: LineWidth(B) = LW
FillValue(B) = Fill: LineValue(B) = LV: Storedtext(B) = St
TextSize(B) = TS: TextFont(B) = TF: TextItalic(B) = TB: TextItalic(B) = TI
FlagAlign(B) = FA: OFFSET(B) = OF
CarriageCount(B) = CR
W(B) = W1: N(B) = N2: E(B) = E3: S(B) = S4
P1X(B) = p1XX: P1Y(B) = p1YY: P2X(B) = p2XX: p2Y(B) = p2YY
h1X(B) = h1XX: h1Y(B) = h1YY: h2X(B) = h2XX: h2Y(B) = h2YY

UpdateObjects

End Sub

Private Sub ToolButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ArtWork.ForeColor = BLACK
    ArtWork.MousePointer = 0

LastToolUsed = CurrentTool: NewTool = Y \ 32
    If Y < 288 Then
    ToolButtons.Line (0, CurrentTool * 32 + 1)-(31, CurrentTool * 32 + 31), , BF
    ToolButtons.Line (0, NewTool * 32 + 1)-(31, NewTool * 32 + 31), , BF
    CurrentTool = NewTool
    End If

Select Case CurrentTool
    Case 1
     If BezHandle1.Visible Then Exit Sub
     If FlagBezier(B) = B Then ShowDotted = True: BezUpdate
    Case 2
     ZoomBar.Visible = True
    Case 6
      LineButtons.Visible = True
    Case 7
      PaintButtons.Visible = True
    Case 8
      ArtWork.MousePointer = 2 'Cross
      '''''''''''''
End Select

If LastToolUsed = TextTool Then ReDrawText
If CurrentTool = 0 And Total > 0 Then
    Nodes
    Handles
    Else
    For R = 0 To 7
    Handle(R).Visible = False
    Next R
    End If

If CurrentTool <> 1 And BezHandle1.Visible Then ShowDotted = False: BezUpdate
If CurrentTool <> 2 Then ZoomBar.Visible = False
If CurrentTool <> 6 Then LineButtons.Visible = False
If CurrentTool <> 7 Then PaintButtons.Visible = False

End Sub

Private Sub ToolButtons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RulerGuides Index, Button, Shift, X, Y

End Sub

Private Sub ToolButtons_Paint()
    
ToolButtons.Line (0, CurrentTool * 32 + 1)-(31, CurrentTool * 32 + 31), , BF
    
End Sub

Private Sub Txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    X = Txt.Left: Y = ArtWork.CurrentY
    ArtWork_MouseUp 2, 0, X, Y
  End If

End Sub

Private Sub UpdateDeskTop()

SW = ((ArtWork.Width) \ 2) * PixelX 'Redraw Page
SH = ((ArtWork.Height) \ 2) * PixelY
OS = 75

    ArtWork.DrawWidth = 1
    PageX1 = SW - (163 * PixelX) * 1 'For large monitor multiply
    PageY1 = SH - (212 * PixelY) * 1
    PageX2 = SW + (162 * PixelX) * 1
    PageY2 = SH + (213 * PixelY) * 1
    ArtWork.Line (PageX1 + OS, PageY1 + OS)-(PageX2 + OS, PageY2 + OS), GRAY, BF
    ArtWork.Line (PageX1, PageY1)-(PageX2, PageY2), WHITE, BF
    ArtWork.Line (PageX1, PageY1)-(PageX2, PageY2), BLACK, B
''''''''''''''''
For R = 1 To TGL 'Redraw Any Guide Lines
    ArtWork.DrawStyle = DOT
    ArtWork.Line (-5000, -5000)-(GuideX(R), 24000), QBColor(1), B
    ArtWork.Line (-5000, -5000)-(24000, GuideY(R)), QBColor(1), B
Next R
X2 = -5000: Y2 = -5000
ArtWork.DrawStyle = SOLID

End Sub

Private Sub UpdateObjects()

LastMove = B
ArtWork.DrawMode = COPY_PEN
ArtWork.DrawStyle = SOLID
ArtWork.AutoRedraw = True

If Draw = True Then GoTo SkipErase

ArtWork.DrawWidth = 20
ArtWork.Line (W(B), N(B))-(E(B), S(B)), WHITE, BF
ArtWork.Line (XX2, YY2)-(XX1, YY1), WHITE, BF: B = 1
ArtWork.DrawWidth = 1
UpdateDeskTop

Do While B < Total + 1
ArtWork.DrawWidth = LineWidth(B)
If LineValue(B) > 0 Then ArtWork.DrawStyle = 5 Else ArtWork.DrawStyle = SOLID
SkipErase:

Select Case B
  Case FlagCircle(B)
        X1 = W(B): Y1 = N(B): X2 = E(B): Y2 = S(B)
        GoodCircle = PerfectCircle(X1, Y1, X2, Y2)
        ArtWork.FillColor = PaintColor(B): ArtWork.FillStyle = FillValue(B)
        ArtWork.Circle (XCenter, YCenter), Radius, OutLineColor(B), , , Aspect
        ArtWork.FillColor = 0: ArtWork.FillStyle = 1
  Case FlagLine(B)
    ArtWork.Line (W(B), N(B))-(E(B), S(B)), OutLineColor(B)
  Case FlagBox(B)
    If FillValue(B) = 1 Then GoTo SkipFill
    ArtWork.Line (W(B), N(B))-(E(B), S(B)), PaintColor(B), BF
SkipFill:
    ArtWork.Line (W(B), N(B))-(E(B), S(B)), OutLineColor(B), B
  Case FlagBezier(B)
    p1.X = P1X(B): p1.Y = P1Y(B): p2.X = P2X(B): p2.Y = p2Y(B)
    h1.X = h1X(B): h1.Y = h1Y(B): h2.X = h2X(B): h2.Y = h2Y(B)
    ArtWork.ForeColor = OutLineColor(B)
    BezDraw
  Case FlagText(B)
    ArtWork.CurrentX = OFFSET(B): ArtWork.CurrentY = N(B)
    SaveScaleLeft = ArtWork.ScaleLeft
    ArtWork.ScaleLeft = -ArtWork.CurrentX + Scaling
    ArtWork.FontName = TextFont(B)
    ArtWork.ForeColor = PaintColor(B)
    ArtWork.FontSize = TextSize(B) * TextPoint
    ArtWork.FontBold = TextBold(B)
    ArtWork.FontItalic = TextItalic(B)
    wFlag = FlagAlign(B)
    TA = SetTextAlign(ArtWork.hDC, wFlag)
    ArtWork.Print Storedtext(B)
    ArtWork.ScaleLeft = SaveScaleLeft
End Select
'''''
If Draw = True Then Draw = False: ArtWork.AutoRedraw = False: B = LastMove: Exit Sub
B = B + 1
Loop

B = LastMove
ArtWork.AutoRedraw = False
ArtWork.DrawWidth = 1: ArtWork.DrawStyle = SOLID

If DeleteObject = True Then Exit Sub
If CurrentTool = 1 Or CurrentTool = 2 Then Exit Sub
If SideRuler = True Or TopRuler = True Then Exit Sub
Handles
Nodes

End Sub

Private Sub VScroll_Change()
 
ArtWork.Enabled = False
Dim WM_MOUSEMOVE%
WM_MOUSEMOVE% = &H200
API = SendMessageByNum(ArtWork.hwnd, WM_MOUSEMOVE%, 0, 0)

 ArtWork.Move ArtWork.Left, -VScroll.value
 Ruler(0).Move Ruler(0).Left, ArtWork.Top - 250
 SSM = (VScroll.value * PixelY) * PerSH

End Sub

Private Sub WinPrint()

On Error GoTo PrintErr:
CMDialog1.CancelError = True
CMDialog1.Flags = &H2& Or &H4& Or &H8& Or &H20&
CMDialog1.Action = 5
NumCopies = CMDialog1.Copies
Screen.MousePointer = 11

For R = 1 To NumCopies

  PrinterScale = Printer.ScaleWidth 'Save scale for fontsize
  Printer.ScaleLeft = PageX1 'Left Margin
  Printer.ScaleTop = PageY1  'Top Margin
  Printer.ScaleWidth = (PageX2 - PageX1) 'Right Enlarge
  Printer.ScaleHeight = (PageY2 - PageY1) 'Bottom Enlarge

B = 1
Do While B < Total + 1
Printer.DrawWidth = LineWidth(B)
If LineValue(B) > 0 Then Printer.DrawStyle = 5 Else Printer.DrawStyle = SOLID
Select Case B
  Case FlagCircle(B)
        X1 = W(B): Y1 = N(B): X2 = E(B): Y2 = S(B)
        GoodCircle = PerfectCircle(X1, Y1, X2, Y2)
        Printer.FillColor = PaintColor(B): Printer.FillStyle = FillValue(B)
        Printer.Circle (XCenter, YCenter), Radius, OutLineColor(B), , , Aspect
        Printer.FillColor = 0: Printer.FillStyle = 1
  Case FlagLine(B)
    Printer.Line (W(B), N(B))-(E(B), S(B)), OutLineColor(B)
  Case FlagBox(B)
    If FillValue(B) = 1 Then GoTo SkipFill2
    Printer.Line (W(B), N(B))-(E(B), S(B)), PaintColor(B), BF
SkipFill2:
    Printer.Line (W(B), N(B))-(E(B), S(B)), OutLineColor(B), B
  Case FlagBezier(B)
    p1.X = P1X(B): p1.Y = P1Y(B): p2.X = P2X(B): p2.Y = p2Y(B)
    h1.X = h1X(B): h1.Y = h1Y(B): h2.X = h2X(B): h2.Y = h2Y(B)
    Printer.ForeColor = OutLineColor(B): FlagPrinter = 2
    BezDraw
  Case FlagText(B)
    Printer.CurrentX = OFFSET(B): Printer.CurrentY = N(B)
    SaveScaleLeft = Printer.ScaleLeft
    Printer.ScaleLeft = -(OFFSET(B) - Printer.ScaleLeft)
    Printer.FontName = TextFont(B)
    PFontScale! = (PrinterScale) / (Printer.ScaleWidth)
    Printer.FontSize = TextSize(B) * PFontScale!
    Printer.FontBold = TextBold(B)
    Printer.FontItalic = TextItalic(B)
    Printer.ForeColor = PaintColor(B)
    wFlag = FlagAlign(B)
    TA = SetTextAlign(Printer.hDC, wFlag)
    Printer.Print Storedtext(B)
    Printer.ScaleLeft = SaveScaleLeft
End Select
'''''
B = B + 1
Loop

Printer.EndDoc
Next R
    Screen.MousePointer = 0
    FlagPrinter = 0
    Exit Sub

PrintErr:
    Screen.MousePointer = 0
    FlagPrinter = 0
    Exit Sub

End Sub

Private Sub ZoomBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

CurrentTool = 2: ToolButtons_Paint: CurrentTool = LastToolUsed: ToolButtons_Paint
CurrentZoom = X \ 32: ZoomBar.Visible = False: N_H_Clear

Select Case CurrentZoom
Case ZoomIn 'First Button
   If Zoom_Out = True Or Zoom_In = True Then Beep: Exit Sub
   ArtWork.AutoRedraw = True: ArtWork.Cls: ArtWork.AutoRedraw = False
   Scaling = 3750: Zoom_In = True: Zoom_Out = False
   Zoom = -48: NodeZoom = -21: TextPoint = 1.7
   Txt.FontSize = Txt.FontSize * TextPoint
   ZoomScale 2, 3750, 3750, -7500, -6375
Case ZoomOut 'Second Button
   If Zoom_Out = True Or Zoom_In = True Then Beep: Exit Sub
   ArtWork.AutoRedraw = True: ArtWork.Cls: ArtWork.AutoRedraw = False
   Scaling = -5000: Zoom_In = False: Zoom_Out = True
   Zoom = 65: NodeZoom = 28: TextPoint = 0.65
   Txt.FontSize = Txt.FontSize * TextPoint
   ZoomScale 2, -5000, -5000, 10000, 8500

Case ZoomAll 'Third Button
 X1 = W(1): Y1 = N(1): X2 = E(1): Y2 = S(1)
 For i = 1 To Total
  If X1 > W(i) Then X1 = W(i)
  If Y1 > N(i) Then Y1 = N(i)
  If X2 < E(i) Then X2 = E(i)
  If Y2 < S(i) Then Y2 = S(i)
 Next i
''''''
 For i = 1 To Total
  If X1 > E(i) Then X1 = E(i)
  If Y1 > S(i) Then Y1 = S(i)
  If X2 < W(i) Then X2 = W(i)
  If Y2 < N(i) Then Y2 = N(i)
Next i
''''''
ArtWork.Line (X1, Y1)-(X2, Y2), QBColor(4), B

Case ZoomPage 'Fourth Button
 ArtWork.AutoRedraw = True: ArtWork.Cls: ArtWork.AutoRedraw = False
 ArtWork.Scale: PerSW = 1: PerSH = 1
 Zoom = 0: NodeZoom = 0: TextPoint = 1
 Scaling = 0: Zoom_In = False: Zoom_Out = False
 Txt.FontSize = 12
   
     For R = 0 To 7
      Handle(R).Scale
     Next R

     For R = 0 To 3
      Node(R).Scale
     Next R

   ArtWork.MousePointer = 1
   VScroll.value = 0: HScroll.value = 0
   ZoomScale 2, 0, 0, 0, 0
End Select '''''''''''''''''''''''''''''

StatusArea_Paint
ToolButtons_MouseDown Button, Shift, X, Y

End Sub

Private Sub ZoomBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RulerGuides Index, Button, Shift, X, Y

End Sub

Private Sub ZoomScale(Index%, ZLeft%, ZTop%, ZWidth%, ZHeight%)

   ArtWork.ScaleLeft = ArtWork.ScaleLeft + ZLeft
   ArtWork.ScaleTop = ArtWork.ScaleTop + ZTop
   ArtWork.ScaleWidth = ArtWork.ScaleWidth + ZWidth
   ArtWork.ScaleHeight = ArtWork.ScaleHeight + ZHeight
   NewSW = ArtWork.ScaleWidth: NewSH = ArtWork.ScaleHeight
   PerSW = NewSW / OldSW: PerSH = NewSH / OldSH
   BSM = BSM * PerSW: SSM = SSM * PerSH
'''''''''''
   Ruler(0).ScaleTop = ArtWork.ScaleTop
   Ruler(0).ScaleHeight = ArtWork.ScaleHeight
   Ruler(1).ScaleLeft = ArtWork.ScaleLeft
   Ruler(1).ScaleWidth = ArtWork.ScaleWidth

  Ruler(0).Refresh
UpdateObjects

End Sub

