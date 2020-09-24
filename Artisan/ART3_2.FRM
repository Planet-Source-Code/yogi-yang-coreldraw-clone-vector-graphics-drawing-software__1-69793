VERSION 5.00
Begin VB.Form NodeEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Edit Nodes"
   ClientHeight    =   855
   ClientLeft      =   3495
   ClientTop       =   3165
   ClientWidth     =   2145
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
   ScaleHeight     =   57
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   143
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton To_Curve 
      Appearance      =   0  'Flat
      Caption         =   "To Curve"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton To_Line 
      Appearance      =   0  'Flat
      Caption         =   "To Line"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "NodeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
 
 NodeEdit.Hide

End Sub

Private Sub Form_Load()

NodeEdit.Left = (Screen.Width - NodeEdit.Width) / 2
NodeEdit.Top = (Screen.Height - NodeEdit.Height) / 2

End Sub

Private Sub Form_Resize()

NodeEdit.Left = (Screen.Width - NodeEdit.Width) / 2
NodeEdit.Top = (Screen.Height - NodeEdit.Height) / 2

End Sub

Private Sub To_Curve_Click()
 
Convert_Curve = True: NodeEdit.Hide
 
End Sub

Private Sub To_Line_Click()

Convert_Line = True: NodeEdit.Hide

End Sub

