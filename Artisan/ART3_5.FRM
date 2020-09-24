VERSION 5.00
Begin VB.Form Credits 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Artisan 1.0"
   ClientHeight    =   5190
   ClientLeft      =   1005
   ClientTop       =   2175
   ClientWidth     =   8370
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "ART3_5.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

  Load Artisan
  Artisan.Show
  Unload Credits

End Sub

Private Sub Form_Load()
    
  Credits.WindowState = 2
    
End Sub

Private Sub Form_Resize()

FW = 10
FontSize = 24
MSG$ = "Artisan is loading"

For R = 1 To 30
ForeColor = QBColor(0)
CurrentX = FW - OFFSET
Print MSG$
Next R

MSG$ = "Artisan"
CurrentY = 1
For R = 1 To 30
ForeColor = QBColor(15)
CurrentX = FW + 3
Print MSG$
Print
Next R

End Sub

