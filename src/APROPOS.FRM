VERSION 5.00
Begin VB.Form APropos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A Propos ..."
   ClientHeight    =   2250
   ClientLeft      =   825
   ClientTop       =   2055
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "APROPOS.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2250
   ScaleMode       =   0  'User
   ScaleWidth      =   5640
   Begin VB.PictureBox Pic_ApplicationIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   255
      Picture         =   "APROPOS.frx":030A
      ScaleHeight     =   487.619
      ScaleMode       =   0  'User
      ScaleWidth      =   487.619
      TabIndex        =   3
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblNomAppl 
      Caption         =   "WinMath"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   990
      TabIndex        =   0
      Top             =   165
      Width           =   1920
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   990
      TabIndex        =   1
      Top             =   735
      Width           =   1470
   End
   Begin VB.Label lblAuteur 
      Caption         =   "Auteur : David BLUM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   990
      TabIndex        =   2
      Top             =   1200
      Width           =   2385
   End
End
Attribute VB_Name = "APropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   ' Center the AboutBox on the screen
   Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub


Private Sub Lbl_Microsoft_Click()

End Sub


