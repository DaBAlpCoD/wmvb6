VERSION 5.00
Begin VB.Form Fenetre_Val 
   AutoRedraw      =   -1  'True
   Caption         =   "Valeurs"
   ClientHeight    =   6990
   ClientLeft      =   7575
   ClientTop       =   1740
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6990
   ScaleWidth      =   3975
   Begin VB.Label lblNote 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblFormule 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Fenetre_Val"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
    Fenetre_Val.ForeColor = Fenetre_Val.BackColor
    Fenetre_Val.PSet (0, 3900), Fenetre_Val.BackColor
    Fenetre_Val.Print Space$(50)
    Fenetre_Val.PSet (0, 4200), Fenetre_Val.BackColor
    Fenetre_Val.Print Space$(50)
    Fenetre_Val.Print Space$(50)
End Sub

