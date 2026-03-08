VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreDefRegreLin 
   Caption         =   "Définition de la droite [régression linéaire]"
   ClientHeight    =   6390
   ClientLeft      =   255
   ClientTop       =   645
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6390
   ScaleWidth      =   9255
   Begin MSFlexGridLib.MSFlexGrid GridPts 
      Height          =   1215
      Left            =   4680
      TabIndex        =   35
      Top             =   4200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2143
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GridPar 
      Height          =   855
      Left            =   120
      TabIndex        =   45
      Top             =   3840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1508
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pctColorPts 
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
      Height          =   255
      Index           =   5
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   60
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   59
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   58
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
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
      Height          =   255
      Index           =   2
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   57
      Top             =   5760
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   5760
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   5760
      Width           =   255
   End
   Begin VB.PictureBox pctColorCou 
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
      Height          =   255
      Index           =   5
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pctColorCou 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pctColorCou 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pctColorCou 
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
      Height          =   255
      Index           =   2
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Top             =   5760
      Width           =   255
   End
   Begin VB.PictureBox pctColorCou 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   5760
      Width           =   255
   End
   Begin VB.PictureBox pctColorCou 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   5760
      Width           =   255
   End
   Begin VB.TextBox txtNbPts 
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
      Left            =   5520
      TabIndex        =   44
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox pctColorPoints 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   5760
      Width           =   495
   End
   Begin VB.CheckBox GrilleTracer 
      Caption         =   "tracer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pctColorCourbe 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox DefYmax 
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
      Left            =   7080
      TabIndex        =   5
      Text            =   "4"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox DefXmax 
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
      Left            =   2280
      TabIndex        =   4
      Text            =   "4"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox DefYmin 
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
      Left            =   7080
      TabIndex        =   3
      Text            =   "-4"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox DefXmin 
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
      Left            =   2280
      TabIndex        =   2
      Text            =   "-4"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox MemeEchelle 
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
      TabIndex        =   1
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox AxesAfficher 
      Caption         =   "afficher les valeurs des graduations"
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
      Left            =   2400
      TabIndex        =   24
      Top             =   960
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox CadreAfficher 
      Caption         =   "afficher les valeurs des graduations"
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
      TabIndex        =   20
      Top             =   960
      Value           =   2  'Grayed
      Width           =   2055
   End
   Begin VB.CheckBox AxesGraduer 
      Caption         =   "graduer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   600
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox CadreGraduer 
      Caption         =   "graduer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox AxesTracer 
      Caption         =   "tracer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   240
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox CadreTracer 
      Caption         =   "tracer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Frame CadreGarder 
      Caption         =   "Tracés précédents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      TabIndex        =   30
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton OptionGarder 
         Caption         =   "garder "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptionEffacer 
         Caption         =   "effacer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label lblRemCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   48
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label lblValCor 
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
      Height          =   255
      Left            =   2760
      TabIndex        =   47
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblCor 
      Caption         =   "Coefficient de corrélation R :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   3480
      Y2              =   5400
   End
   Begin VB.Label lblMemeEch 
      Caption         =   "même échelle sur les deux axes"
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
      Left            =   360
      TabIndex        =   40
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblNbPts 
      Caption         =   "Nombre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   43
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblPts 
      Caption         =   "Points par où on cherche à faire passer la droite"
      Height          =   255
      Left            =   4680
      TabIndex        =   42
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label lblCoorPts 
      Caption         =   "Coordonnées :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblChoixCoulPts 
      Caption         =   "Couleur choisie :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   38
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label EtiquetteCoulPoints 
      Caption         =   "Couleur des points :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   37
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblPar 
      Caption         =   "Paramètres entrant dans la définition de la droite"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label EtiquetteGrille 
      Caption         =   "Grille :"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblChoixCoulCou 
      Caption         =   "Couleur choisie :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   29
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label EtiquetteFactE 
      Caption         =   "FE = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   34
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label EtiquetteCoulDroite 
      Caption         =   "Couleur de la droite :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "f(X) = A*X + B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   26
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblDefiniDroi 
      Caption         =   "Fonction définissant la droite :"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label EtiquetteYmax 
      Caption         =   "Ymax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label EtiquetteXmax 
      Caption         =   "Xmax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label EtiquetteYmin 
      Caption         =   "Ymin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label EtiquetteLimitesY 
      Caption         =   "Limites de l'axe des ordonnées "
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
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label EtiquetteXmin 
      Caption         =   "Xmin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label EtiquetteLimitesX 
      Caption         =   "Limites de l'axe des abscisses :"
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
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label EtiquetteAccolY 
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label EtiquetteAccolX 
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label EtiquetteFE 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"DEFRELIN.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Label EtiquetteAccolRem 
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   30
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1920
      TabIndex        =   27
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label EtiquetteAxes 
      Caption         =   "Axes :"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   0
      Width           =   735
   End
   Begin VB.Label EtiquetteCadre 
      Caption         =   "Cadre :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
         Begin VB.Menu mnuOuvImage 
            Caption         =   "Une &image..."
         End
         Begin VB.Menu mnuOuvPoints 
            Caption         =   "Un &ensemble de points..."
         End
      End
      Begin VB.Menu mnuEnregistrer 
         Caption         =   "En&registrer"
         Begin VB.Menu mnuEnrImage 
            Caption         =   "Une i&mage..."
         End
         Begin VB.Menu mnuEnrPoints 
            Caption         =   "Un e&nsemble de points..."
         End
      End
   End
   Begin VB.Menu mnuDefPts 
      Caption         =   "&Définir un ensemble de points"
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer"
   End
   Begin VB.Menu mnuTracer 
      Caption         =   "&Tracer..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreDefRegreLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub AxesAfficher_Click()
    If Grax% = 1 Then
        Vax% = AxesAfficher.Value
    Else
        AxesAfficher.Value = 2
    End If
End Sub

Private Sub AxesGraduer_Click()
    Grax% = AxesGraduer.Value
    If Grax% = 0 Then
        AxesAfficher.Value = 2
    ElseIf Grax% = 1 Then
        AxesAfficher.Value = 0
    End If
End Sub

Private Sub AxesTracer_Click()
    Trax% = AxesTracer.Value
End Sub

Private Sub CadreAfficher_Click()
    If Grac% = 1 Then
        Vac% = CadreAfficher.Value
    Else
        CadreAfficher.Value = 2
    End If
End Sub

Private Sub CadreGraduer_Click()
    Grac% = CadreGraduer.Value
    If Grac% = 0 Then
        CadreAfficher.Value = 2
    ElseIf Grac% = 1 Then
        CadreAfficher.Value = 0
    End If
End Sub

Private Sub CadreTracer_Click()
    Trac% = CadreTracer.Value
End Sub



Private Sub DefXmax_Change()
    Select Case Gap%
        Case 1
            On Error Resume Next
            Xmax = CSng(DefXmax.Text)
            If Ech% = 1 Then
                AX = Xmax - Xmin
                AY = FE * AX
                Ymax = AY + Ymin
                DefYmax.Text = Format(Ymax, "0.0000")
            End If
            On Error GoTo 0
        Case 2
            DefXmax.Text = Format(Xmax, "0.0000")
    End Select
End Sub
Private Sub DefXmin_Change()
    Select Case Gap%
        Case 1
            On Error Resume Next
            Xmin = CSng(DefXmin.Text)
            If Ech% = 1 Then
                AX = Xmax - Xmin
                AY = FE * AX
                Ymax = AY + Ymin
                DefYmax.Text = Format(Ymax, "0.0000")
            End If
            On Error GoTo 0
        Case 2
            DefXmin.Text = Format(Xmin, "0.0000")
    End Select
End Sub

Private Sub DefYmax_Change()
    Select Case Gap%
        Case 1
            On Error Resume Next
            Ymax = CSng(DefYmax.Text)
            If Ech% = 1 Then
                AX = Xmax - Xmin
                AY = FE * AX
                Ymin = Ymax - AY
                DefYmin.Text = Format(Ymin, "0.0000")
            End If
            On Error GoTo 0
        Case 2
            DefYmax.Text = Format(Ymax, "0.0000")
    End Select
End Sub


Private Sub DefYmin_Change()
    Select Case Gap%
        Case 1
            On Error Resume Next
            Ymin = CSng(DefYmin.Text)
            If Ech% = 1 Then
                AX = Xmax - Xmin
                AY = FE * AX
                Ymax = AY + Ymin
                DefYmax.Text = Format(Ymax, "0.0000")
            End If
            On Error GoTo 0
        Case 2
            DefYmin.Text = Format(Ymin, "0.0000")
    End Select
End Sub

Private Sub Form_Activate()
    If Gap% = 2 Then
      OptionGarder.Value = True
    End If
    CadreTracer.Value = Trac%
    CadreGraduer.Value = Grac%
    CadreAfficher.Value = Vac%
    AxesTracer.Value = Trax%
    AxesGraduer.Value = Grax%
    AxesAfficher.Value = Vax%
    MemeEchelle.Value = Ech%
    EtiquetteFE.Caption = "     " + Format(FE, "0.000000")
    DefXmin.Text = Format(Xmin, "0.000000")
    DefXmax.Text = Format(Xmax, "0.000000")
    DefYmin.Text = Format(Ymin, "0.000000")
    DefYmax.Text = Format(Ymax, "0.000000")
   '***************************************
   '*****       actualisation ...     *****
   '***************************************
   FenetreDefRegreLin.lblValCor.Caption = ""
   '***************************************
   '***** actualisation des 2 grilles *****
   '***** 1) grille des points ************
   '***************************************
   FenetreDefRegreLin.txtNbPts.Text = Format(NbPtsReLin%, "0")
   For i% = 1 To NbPtsReLin%
      FenetreDefRegreLin.GridPts.Row = i%
      FenetreDefRegreLin.GridPts.Col = 1
      FenetreDefRegreLin.GridPts.Text = Format(Xrl(i%), "0.0000")
      FenetreDefRegreLin.GridPts.Col = 2
      FenetreDefRegreLin.GridPts.Text = Format(Yrl(i%), "0.0000")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefRegreLin.GridPar.Col = 1
   FenetreDefRegreLin.GridPar.Row = 1
   FenetreDefRegreLin.GridPar.Text = ""
   FenetreDefRegreLin.GridPar.Row = 2
   FenetreDefRegreLin.GridPar.Text = ""
   ' ********************************
End Sub

Private Sub Form_Deactivate()
   On Error GoTo Traite_ErreursFoDeReLin
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Xrl, Yrl
   ReDim Xrl(1 To NbPtsReLin%), Yrl(1 To NbPtsReLin%)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsReLin%
      FenetreDefRegreLin.GridPts.Row = i%
      FenetreDefRegreLin.GridPts.Col = 1
      Xrl(i%) = CSng(FenetreDefRegreLin.GridPts.Text)
      FenetreDefRegreLin.GridPts.Col = 2
      Yrl(i%) = CSng(FenetreDefRegreLin.GridPts.Text)
   Next i%
   ' ********************************************
   Exit Sub
Traite_ErreursFoDeReLin:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub Form_Load()
   '-----------------------------------------------------------------
   '-----------        Valeurs par défaut             ---------------
   '-----------      de certaines variables           ---------------
   '-----------------------------------------------------------------
   DoFlag = False
   Coor% = 1
   Expr% = 1
   Gap% = 1
   Trac% = 1
   Grac% = 0
   Vac% = 2
   Trax% = 1
   Grax% = 1
   Vax% = 1
   Ech% = 1
   Xmin = -4
   Xmax = 4
   Ymin = -4
   Ymax = 4
   CouleurCou& = ROUGE
   CouleurPts& = BLEU
   '-----------------------------------------------------------------
   '-----------        Valeurs par défaut             ---------------
   '-----------        d'autres variables             ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   AX = Xmax - Xmin
   AY = Ymax - Ymin
   MX = LX% / AX
   MY = -LY% / AY
   '-----------------------------------------------------
   '---------------- valeurs par défaut -----------------
   '----------------    pour RegreLin   -----------------
   '-----------------------------------------------------
   NbPtsReLin% = 20    ' nombre de points
   ' nombre de paramètres = 2
   ReDim Xrl(1 To NbPtsReLin%), Yrl(1 To NbPtsReLin%)
   ' coeficients de la droite
   Adroite = 0.8
   Bdroite = 0.57
   '************ points sur y = 0.8 x + 0.6 **************
   Xrl(1) = -3.68
   Yrl(1) = -2.44
   Xrl(2) = -3.62
   Yrl(2) = -2.04
   Xrl(3) = -3.18
   Yrl(3) = -2.09
   Xrl(4) = -3.02
   Yrl(4) = -1.75
   Xrl(5) = -2.64
   Yrl(5) = -1.62
   Xrl(6) = -2.29
   Yrl(6) = -1.42
   Xrl(7) = -2.2
   Yrl(7) = -1
   Xrl(8) = -1.68
   Yrl(8) = -0.91
   Xrl(9) = -1.6
   Yrl(9) = -0.64
   Xrl(10) = -1.04
   Yrl(10) = -0.33
   Xrl(11) = -0.68
   Yrl(11) = -0.11
   Xrl(12) = -0.49
   Yrl(12) = 0.31
   Xrl(13) = -0.07
   Yrl(13) = 0.33
   Xrl(14) = 0.51
   Yrl(14) = 0.89
   Xrl(15) = 0.73
   Yrl(15) = 1.33
   Xrl(16) = 1.17
   Yrl(16) = 1.42
   Xrl(17) = 1.47
   Yrl(17) = 1.6
   Xrl(18) = 1.57
   Yrl(18) = 2.06
   Xrl(19) = 2.46
   Yrl(19) = 2.42
   Xrl(20) = 2.97
   Yrl(20) = 3.15
   '***************************************
   RemCor$ = "Remarque :    0 < R < 1" & Chr$(13)
   RemCor$ = RemCor$ + "     si R -> 0 : pas de corrélation"
   RemCor$ = RemCor$ + Chr$(13)
   RemCor$ = RemCor$ + "     si R -> 1 : trés bonne corrélation"
   FenetreDefRegreLin.lblRemCor.Caption = RemCor$
   '***************************************
   '***** mise en place des 2 grilles *****
   '***** 1) grille des points ************
   ' **************************************
   FenetreDefRegreLin.txtNbPts.Text = Format(NbPtsReLin%, "0")
   FenetreDefRegreLin.GridPts.Rows = NbPtsReLin% + 1
   FenetreDefRegreLin.GridPts.Cols = 3
   FenetreDefRegreLin.GridPts.FixedAlignment(0) = 2
   FenetreDefRegreLin.GridPts.FixedAlignment(1) = 2
   FenetreDefRegreLin.GridPts.FixedAlignment(2) = 2
   FenetreDefRegreLin.GridPts.ColWidth(1) = 1500
   FenetreDefRegreLin.GridPts.ColWidth(2) = 1500
   FenetreDefRegreLin.GridPts.Row = 0
   FenetreDefRegreLin.GridPts.Col = 1
   FenetreDefRegreLin.GridPts.Text = "X"
   FenetreDefRegreLin.GridPts.Col = 2
   FenetreDefRegreLin.GridPts.Text = "Y"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefRegreLin.GridPts.Col = 0
   For i% = 1 To NbPtsReLin%
      FenetreDefRegreLin.GridPts.Row = i%
      FenetreDefRegreLin.GridPts.Text = Format(i%, "0")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefRegreLin.GridPar.Rows = 3
   FenetreDefRegreLin.GridPar.Cols = 2
   FenetreDefRegreLin.GridPar.FixedAlignment(0) = 2
   FenetreDefRegreLin.GridPar.FixedAlignment(1) = 2
   FenetreDefRegreLin.GridPar.ColWidth(1) = 2500
   ' ********************************
   ' première colonne
   ' ********************************
   FenetreDefRegreLin.GridPar.Col = 0
   FenetreDefRegreLin.GridPar.Row = 1
   FenetreDefRegreLin.GridPar.Text = "A"
   FenetreDefRegreLin.GridPar.Row = 2
   FenetreDefRegreLin.GridPar.Text = "B"
   ' ********************************
End Sub

Private Sub gridPar_KeyPress(KeyAscii As Integer)
   ParText$ = GridPar.Text
   Select Case KeyAscii
   Case 32 To 168
      ParCar$ = Chr(KeyAscii)
      ParText$ = ParText$ & ParCar$
      GridPar.Text = ParText$
   Case 8
      If Len(GridPar.Text) > 0 Then
         ParText$ = Left$(ParText$, Len(ParText$) - 1)
         GridPar.Text = ParText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub gridPts_KeyPress(KeyAscii As Integer)
   FenetreDefRegreLin.lblValCor.Caption = ""
   PtsText$ = GridPts.Text
   Select Case KeyAscii
   Case 32 To 168
      PtsCar$ = Chr(KeyAscii)
      PtsText$ = PtsText$ & PtsCar$
      GridPts.Text = PtsText$
   Case 8
      If Len(GridPts.Text) > 0 Then
         PtsText$ = Left$(PtsText$, Len(PtsText$) - 1)
         GridPts.Text = PtsText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub GrilleTracer_Click()
    Tric% = GrilleTracer.Value
End Sub

Private Sub MemeEchelle_Click()
    Ech% = MemeEchelle.Value
    If Ech% = 1 Then
        Ymax = FE * (Xmax - Xmin) + Ymin
        DefYmax.Text = Format(Ymax, "0.000000")
    End If
End Sub


Private Sub mnuCalculer_Click()
   Call CalculeRegreLin
End Sub



Private Sub mnuEnrImage_Click()
   On Error GoTo Traite_ErreursEnrIm
   Maths.ctrlCMDialog.DefaultExt = "bmp"
   Maths.ctrlCMDialog.Filter = "Image (*.bmp)|*.bmp"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   SavePicture Dessin.Image, Maths.ctrlCMDialog.FileName
   Exit Sub
Traite_ErreursEnrIm:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuEnrPoints_Click()
   On Error GoTo Traite_ErreursEnrpts
   Maths.ctrlCMDialog.DefaultExt = "pts"
   Maths.ctrlCMDialog.Filter = "Points (*.xy)|*.xy"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   '-----------------------------------------------------
   ' Création du fichier de points
   ' et écriture de leurs coordonnées Xrlt() et Yrl()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbPtsReLin%
   For iloc% = 1 To NbPtsReLin%
      Write #1, Xrl(iloc%)
      Write #1, Yrl(iloc%)
   Next iloc%
   Close #1
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnrpts:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuOuvImage_Click()
   On Error GoTo Traite_ErreursOuvIm
   Maths.ctrlCMDialog.Filter = "Image (*.bmp)|*.bmp"
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   Dessin.Picture = LoadPicture(Maths.ctrlCMDialog.FileName)
   Fenetre_Val.lblFormule.Caption = ""
   Exit Sub
Traite_ErreursOuvIm:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuOuvPoints_Click()
   On Error GoTo Traite_ErreursOuvPts
   Maths.ctrlCMDialog.Filter = "Points (*.xy)|*.xy"
   ' nom de fichier et chemin doivent exister
   ' sinon apparait un message d'erreur spécifique
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   '-----------------------------------------------------
   ' Ouverture et lecture du fichier de points
   ' et écriture de leurs coordonnées dans Xpt() et Ypt()
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, NbPtsReLin%
   ReDim Xrl(1 To NbPtsReLin%)
   ReDim Yrl(1 To NbPtsReLin%)
   For iloc% = 1 To NbPtsReLin%
    Input #1, Xrl(iloc%)
    Input #1, Yrl(iloc%)
   Next iloc%
   Close #1
   '*******************************************************
   '***** actualisation de la grille des points ************
   '********************************************************
   FenetreDefRegreLin.txtNbPts.Text = Format(NbPtsReLin%, "0")
   For i% = 1 To NbPtsReLin%
      FenetreDefRegreLin.GridPts.Row = i%
      FenetreDefRegreLin.GridPts.Col = 1
      FenetreDefRegreLin.GridPts.Text = Format(Xrl(i%), "0.0000")
      FenetreDefRegreLin.GridPts.Col = 2
      FenetreDefRegreLin.GridPts.Text = Format(Yrl(i%), "0.0000")
   Next i%
   ' **************************************
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursOuvPts:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuQuitter_Click()
   FenetreDefRegreLin.Hide
   Fenetre_Prin.Hide
   Fenetre_Val.Hide
   Dessin.Hide
End Sub

Private Sub mnuTracer_Click()
   '-----------------------------------------------------
   '---------   Vérifications avant   -------------------
   '--------- fermeture de la fenêtre -------------------
   '-----------------------------------------------------
   '-----  1- Bornes écrites correctement  --------------
   On Error Resume Next
   Xmax = CSng(DefXmax.Text)
   If Err.Number <> 0 Then
      MsgBox "Xmax est incorrect", 48, "COURBE"
      Exit Sub
   End If
   Xmin = CSng(DefXmin.Text)
   If Err.Number <> 0 Then
      MsgBox "Xmin est incorrect", 48, "COURBE"
      Exit Sub
   End If
   Ymax = CSng(DefYmax.Text)
   If Err.Number <> 0 Then
      MsgBox "Ymax est incorrect", 48, "COURBE"
      Exit Sub
   End If
   Ymin = CSng(DefYmin.Text)
   If Err.Number <> 0 Then
      MsgBox "Ymin est incorrect", 48, "COURBE"
      Exit Sub
   End If
   On Error GoTo 0
   '-----  2- Bornes dans le bon ordre  --------------
   AX = Xmax - Xmin
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
   End If
   If NbPtsReLin% < 2 Then
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "FenetreDefRegreLin.mnuTracer"
      Exit Sub
   ElseIf AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "FenetreDefRegreLin.mnuTracer"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "FenetreDefRegreLin.mnuTracer"
      Exit Sub
   End If
   '-----------------------------------------
   '--------   si quelque chose   -----------
   '--------    a été modifié,    -----------
   '--------   calculer d'abord   -----------
   '-----------------------------------------
   If FenetreDefRegreLin.lblValCor.Caption = "" Then
      mnuCalculer_Click
   End If
   '-----------------------------------------------------
   '--------         Ouverture           ----------------
   '--------  des fenêtres pour le tracé ----------------
   '-----------------------------------------------------
   Fenetre_Prin.Show
   Fenetre_Val.Show
   Dessin.Show
   '--------------------------------------------------------------
   '--------- Affichage de la formule ----------------------------
   '---------     définissant la courbe        -------------------
   '--------------------------------------------------------------
   Fenetre_Prin.Caption = "Régression linéaire"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Détermination par la méthode des moindres carrés des paramètres d'une droite pour qu'elle soit le mieux ajustée à un nuage de points."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   TextVar2$ = Format(Adroite, "0.000000") & " X + " & Format(Bdroite, "0.000000")
   Fenetre_Val.lblFormule.Caption = "F(X) = " & TextVar2$
   '--------------------------------------------------------------
   '--------- Affichage du coefficient de corrélation R ----------
   '--------------------------------------------------------------
   Fenetre_Val.lblNote.ForeColor = CouleurCou&
   Fenetre_Val.lblNote.Caption = " Coefficient de corrélation :" & Chr$(13) & " R = " & Format(Rrl, "0.0000")
   '--------------------------------------------------------------
End Sub

Private Sub OptionEffacer_Click()
    Gap% = 1
End Sub

Private Sub OptionGarder_Click()
    Gap% = 2
End Sub

Private Sub pctColorCou_Click(Index As Integer)
   pctColorCourbe.BackColor = pctColorCou(Index).BackColor
   CouleurCou& = pctColorCourbe.BackColor
End Sub

Private Sub pctColorPts_Click(Index As Integer)
   pctColorPoints.BackColor = pctColorPts(Index).BackColor
   CouleurPts& = pctColorPoints.BackColor
End Sub

Private Sub RegreLin()
   On Error GoTo Traite_ErreursReLin
   '******************************************************
   '               Initialisations
   '******************************************************
   XG = 0
   YG = 0
   XX = 0
   XY = 0
   YY = 0
   '******************************************************
   '                   Calculs
   '******************************************************
   For i% = 1 To NbPtsReLin%
      Xloc = Xrl(i%)
      Yloc = Yrl(i%)
      XG = XG + Xloc
      YG = YG + Yloc
      XX = XX + Xloc * Xloc
      XY = XY + Xloc * Yloc
      YY = YY + Yloc * Yloc
   Next i%
   '******************************************************
   '               Moyennes :
   '******************************************************
   '------------------------------------------------------
   '               Centre de gravité ( X et Y moyens )
   '------------------------------------------------------
   '  X moyen :
   XG = XG / NbPtsReLin%
   '  Y moyen :
   YG = YG / NbPtsReLin%
   '------------------------------------------------------
   '               Autres moyennes
   '------------------------------------------------------
   '  X² moyen :
   XX = XX / NbPtsReLin%
   '  Y² moyen :
   YY = YY / NbPtsReLin%
   '  XY moyen :
   XY = XY / NbPtsReLin%
   '------------------------------------------------------
   '******************************************************
   '               Variances :
   '******************************************************
   SX = XX - XG * XG
   SY = YY - YG * YG
   SS = XY - XG * YG
   '******************************************************
   '               A et B
   '******************************************************
   If SX = 0 Then
      Adroite = 100
   Else
      Adroite = SS / SX
   End If
   Bdroite = YG - Adroite * XG
   '******************************************************
   '               Coefficient de corrélation R
   '******************************************************
   If SS = 0 Then
      Rrl = 1
   Else
      Rrl = Abs(SS) / Sqr(SX) / Sqr(SY)
   End If
   '******************************************************
   Exit Sub
Traite_ErreursReLin:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub txtNbPts_Change()
   NbPtsReLin% = CInt(txtNbPts.Text)
   GridPts.Rows = NbPtsReLin% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   GridPts.Col = 0
   For i% = 1 To NbPtsReLin%
      GridPts.Row = i%
      GridPts.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub


Public Sub CalculeRegreLin()
   '--------------------------------------------------------------
   If NbPtsReLin% < 2 Then
      Beep
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "REGRELIN"
      Exit Sub
   End If
   '--------------------------------------------------------------
   Erase Xrl, Yrl
   ReDim Xrl(1 To NbPtsReLin%), Yrl(1 To NbPtsReLin%)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsReLin%
      FenetreDefRegreLin.GridPts.Row = i%
      FenetreDefRegreLin.GridPts.Col = 1
      Xrl(i%) = CSng(FenetreDefRegreLin.GridPts.Text)
      FenetreDefRegreLin.GridPts.Col = 2
      Yrl(i%) = CSng(FenetreDefRegreLin.GridPts.Text)
   Next i%
   ' ********************************************
   Call RegreLin
   '--------------------------------------------------------------
   '------ Affichage des paramètres de la droite -----------------
   '---------------- après l'affinement --------------------------
   '--------------------------------------------------------------
   FenetreDefRegreLin.GridPar.Col = 1
   FenetreDefRegreLin.GridPar.Row = 1
   FenetreDefRegreLin.GridPar.Text = Format(Adroite, "0.000000")
   FenetreDefRegreLin.GridPar.Row = 2
   FenetreDefRegreLin.GridPar.Text = Format(Bdroite, "0.000000")
   '--------------------------------------------------------------
   '----------- Affichage du coefficient de corrélation  Rrl -----
   '--------------------------------------------------------------
   lblValCor.Caption = Format(Rrl, "0.000000")
   '--------------------------------------------------------------
End Sub
