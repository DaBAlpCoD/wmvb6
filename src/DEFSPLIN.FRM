VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreDefSpline 
   Caption         =   "Définition de la courbe [Spline]"
   ClientHeight    =   6885
   ClientLeft      =   210
   ClientTop       =   825
   ClientWidth     =   10320
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
   ScaleHeight     =   6885
   ScaleWidth      =   10320
   Begin MSFlexGridLib.MSFlexGrid GridPts 
      Height          =   1215
      Left            =   4680
      TabIndex        =   45
      Top             =   4080
      Width           =   4935
      _ExtentX        =   8705
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
      Height          =   1215
      Left            =   120
      TabIndex        =   46
      Top             =   4080
      Width           =   4335
      _ExtentX        =   7646
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
      Left            =   5400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   5760
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
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   58
      Top             =   5760
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
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   57
      Top             =   5760
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
      Left            =   5400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   5520
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
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   5520
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
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   5520
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
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Top             =   5760
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
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   5760
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
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Top             =   5760
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
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   5520
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
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   5520
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
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   48
      Top             =   5520
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
      Left            =   6000
      TabIndex        =   44
      Top             =   3720
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
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   5520
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
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   5520
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
      Left            =   7200
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
      Left            =   7200
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
   Begin VB.Label lblCalcul 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7200
      TabIndex        =   59
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"DEFSPLIN.frx":0000
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
      TabIndex        =   47
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   3480
      Y2              =   5280
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
      Caption         =   "Nombre :  n="
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
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblPts 
      Caption         =   "Points par où doit passer la courbe"
      Height          =   255
      Left            =   4680
      TabIndex        =   42
      Top             =   3480
      Width           =   3135
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
      Left            =   7080
      TabIndex        =   41
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblValPar 
      Caption         =   "Valeurs :"
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
      TabIndex        =   35
      Top             =   3720
      Width           =   975
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
      Left            =   5760
      TabIndex        =   38
      Top             =   5520
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
      Height          =   375
      Left            =   3720
      TabIndex        =   37
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblPar 
      Caption         =   "Paramètres entrant dans la définition de la courbe"
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
      Height          =   375
      Left            =   2040
      TabIndex        =   29
      Top             =   5520
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
   Begin VB.Label EtiquetteCoulCourbe 
      Caption         =   "Couleur de la courbe :"
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
      TabIndex        =   33
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "f(X) = Union des A(I) X^3  + B(I) X^2 + C(I) X + D(I)"
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
      Width           =   4935
   End
   Begin VB.Label lblDefiniCour 
      Caption         =   "Fonction définissant la courbe :"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   3015
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
      Left            =   6600
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
      Left            =   6600
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
      Left            =   5040
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
      Left            =   6360
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
         Caption         =   "En&registrer..."
         Begin VB.Menu mnuEnrImage 
            Caption         =   "Une i&mage..."
         End
         Begin VB.Menu mnuEnrPoints 
            Caption         =   "Un e&nsemble de points..."
         End
      End
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
Attribute VB_Name = "FenetreDefSpline"
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
   '***** actualisation des 2 grilles *****
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   FenetreDefSpline.txtNbPts.Text = Format(NbPtsSp%, "0")
   For i% = 1 To NbPtsSp%
      FenetreDefSpline.GridPts.Row = i%
      FenetreDefSpline.GridPts.Col = 1
      FenetreDefSpline.GridPts.Text = Format(Xsp(i%), "0.000")
      FenetreDefSpline.GridPts.Col = 2
      FenetreDefSpline.GridPts.Text = Format(Ysp(i%), "0.000")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefSpline.GridPar.Col = 1
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Aspline(i%), "0.000")
   Next i%
   FenetreDefSpline.GridPar.Col = 2
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Bspline(i%), "0.000")
   Next i%
   FenetreDefSpline.GridPar.Col = 3
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Cspline(i%), "0.000")
   Next i%
   FenetreDefSpline.GridPar.Col = 4
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Dspline(i%), "0.000")
   Next i%
   ' **************************************
End Sub

Private Sub Form_Deactivate()
   On Error GoTo Traite_ErreursFoDeSpline
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Xsp, Ysp
   ReDim Xsp(1 To NbPtsSp%), Ysp(1 To NbPtsSp%)
   ReDim Aspline(1 To NbPtsSp% - 1)
   ReDim Bspline(1 To NbPtsSp% - 1)
   ReDim Cspline(1 To NbPtsSp% - 1)
   ReDim Dspline(1 To NbPtsSp% - 1)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsSp%
      FenetreDefSpline.GridPts.Row = i%
      FenetreDefSpline.GridPts.Col = 1
      If FenetreDefSpline.GridPts.Text = "" Then
         FenetreDefSpline.GridPts.Text = "0"
         Xsp(i%) = 0
      Else
         Xsp(i%) = CSng(FenetreDefSpline.GridPts.Text)
      End If
      FenetreDefSpline.GridPts.Col = 2
      If FenetreDefSpline.GridPts.Text = "" Then
         FenetreDefSpline.GridPts.Text = "0"
         Ysp(i%) = 0
      Else
         Ysp(i%) = CSng(FenetreDefSpline.GridPts.Text)
      End If
   Next i%
   ' ************************
   ' mise en ordre des points
   ' suivant les X croissants
   ' ************************
   Call OrdonnePoints(NbPtsSp%, Xsp(), Ysp())
   ' ********************************************
   ' affectation de leurs valeurs aux paramètres
   ' ********************************************
   FenetreDefSpline.GridPar.Col = 1
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      If FenetreDefSpline.GridPar.Text = "" Then
         FenetreDefSpline.GridPar.Text = "0"
         Aspline(i%) = 0
      Else
         Aspline(i%) = CSng(FenetreDefSpline.GridPar.Text)
      End If
   Next i%
   FenetreDefSpline.GridPar.Col = 2
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      If FenetreDefSpline.GridPar.Text = "" Then
         FenetreDefSpline.GridPar.Text = "0"
         Bspline(i%) = 0
      Else
         Bspline(i%) = CSng(FenetreDefSpline.GridPar.Text)
      End If
   Next i%
   FenetreDefSpline.GridPar.Col = 3
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      If FenetreDefSpline.GridPar.Text = "" Then
         FenetreDefSpline.GridPar.Text = "0"
         Cspline(i%) = 0
      Else
         Cspline(i%) = CSng(FenetreDefSpline.GridPar.Text)
      End If
   Next i%
   FenetreDefSpline.GridPar.Col = 4
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      If FenetreDefSpline.GridPar.Text = "" Then
         FenetreDefSpline.GridPar.Text = "0"
         Dspline(i%) = 0
      Else
         Dspline(i%) = CSng(FenetreDefSpline.GridPar.Text)
      End If
   Next i%
   ' **************************************
   Exit Sub
Traite_ErreursFoDeSpline:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48, "FenetreDefSpline.Deactivate"
      Case Else
         MsgBox Error$, 48, "FenetreDefSpline.Deactivate"
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
   '----------------     pour Spline    -----------------
   '-----------------------------------------------------
   lblCalcul.Caption = "Calcul non effectué"
   NbPtsSp% = 7    ' nombre de points
   ReDim Xsp(1 To NbPtsSp%), Ysp(1 To NbPtsSp%)
   ReDim Aspline(1 To NbPtsSp% - 1)
   ReDim Bspline(1 To NbPtsSp% - 1)
   ReDim Cspline(1 To NbPtsSp% - 1)
   ReDim Dspline(1 To NbPtsSp% - 1)
   '******************************
   ' points sur y = .02x^6-.02x^5-.04x^4+.37x^3+2.18x^2-1.48x-1.56
   Xsp(1) = -3.5
   Ysp(1) = 3
   Xsp(2) = -2.8
   Ysp(2) = 0.5
   Xsp(3) = -1.2
   Ysp(3) = 2
   Xsp(4) = 0.7
   Ysp(4) = -1.5
   Xsp(5) = 1.8
   Ysp(5) = 1
   Xsp(6) = 2.5
   Ysp(6) = 1
   Xsp(7) = 3.2
   Ysp(7) = -0.5
   '***************************************
   '***** mise en place des 2 grilles *****
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   FenetreDefSpline.txtNbPts.Text = Format(NbPtsSp%, "0")
   FenetreDefSpline.GridPts.Cols = 3
   FenetreDefSpline.GridPts.Rows = NbPtsSp% + 1
   FenetreDefSpline.GridPts.FixedAlignment(0) = 2
   FenetreDefSpline.GridPts.FixedAlignment(1) = 2
   FenetreDefSpline.GridPts.FixedAlignment(2) = 2
   FenetreDefSpline.GridPts.ColWidth(0) = 500
   FenetreDefSpline.GridPts.ColWidth(1) = 1500
   FenetreDefSpline.GridPts.ColWidth(2) = 1500
   FenetreDefSpline.GridPts.Row = 0
   FenetreDefSpline.GridPts.Col = 1
   FenetreDefSpline.GridPts.Text = "X"
   FenetreDefSpline.GridPts.Col = 2
   FenetreDefSpline.GridPts.Text = "Y"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefSpline.GridPts.Col = 0
   For i% = 1 To NbPtsSp%
      FenetreDefSpline.GridPts.Row = i%
      FenetreDefSpline.GridPts.Text = Format(i%, "0")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefSpline.GridPar.Cols = 5
   FenetreDefSpline.GridPar.Rows = NbPtsSp%
   FenetreDefSpline.GridPar.FixedAlignment(0) = 2
   FenetreDefSpline.GridPar.ColWidth(0) = 500
   For i% = 1 To 4
      FenetreDefSpline.GridPar.FixedAlignment(i%) = 2
      FenetreDefSpline.GridPar.ColWidth(i%) = 800
   Next i%
   FenetreDefSpline.GridPar.Row = 0
   FenetreDefSpline.GridPar.Col = 1
   FenetreDefSpline.GridPar.Text = "A"
   FenetreDefSpline.GridPar.Col = 2
   FenetreDefSpline.GridPar.Text = "B"
   FenetreDefSpline.GridPar.Col = 3
   FenetreDefSpline.GridPar.Text = "C"
   FenetreDefSpline.GridPar.Col = 4
   FenetreDefSpline.GridPar.Text = "D"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefSpline.GridPar.Col = 0
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub

Private Sub gridPts_KeyPress(KeyAscii As Integer)
   lblCalcul.Caption = "Calcul non effectué"
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
   If lblCalcul.Caption = "Calcul non effectué" Then
      On Error GoTo Traite_ErreursCalSpline
      ' ********************************************
      '               ReDims
      ' ********************************************
      Erase Xsp, Ysp
      ReDim Xsp(1 To NbPtsSp%), Ysp(1 To NbPtsSp%)
      ReDim Aspline(1 To NbPtsSp% - 1)
      ReDim Bspline(1 To NbPtsSp% - 1)
      ReDim Cspline(1 To NbPtsSp% - 1)
      ReDim Dspline(1 To NbPtsSp% - 1)
      ' ********************************************
      ' affectation de leurs valeurs aux coordonnées
      ' ********************************************
      For i% = 1 To NbPtsSp%
         FenetreDefSpline.GridPts.Row = i%
         FenetreDefSpline.GridPts.Col = 1
         If FenetreDefSpline.GridPts.Text = "" Then
            FenetreDefSpline.GridPts.Text = "0"
            Xsp(i%) = 0
         Else
            Xsp(i%) = CSng(FenetreDefSpline.GridPts.Text)
         End If
         FenetreDefSpline.GridPts.Col = 2
         If FenetreDefSpline.GridPts.Text = "" Then
            FenetreDefSpline.GridPts.Text = "0"
            Ysp(i%) = 0
         Else
            Ysp(i%) = CSng(FenetreDefSpline.GridPts.Text)
         End If
      Next i%
      ' ************************
      ' mise en ordre des points
      ' suivant les X croissants
      ' ************************
      Call OrdonnePoints(NbPtsSp%, Xsp(), Ysp())
      ' ******************************
      ' calcul des paramètres A,B,C,D
      ' ******************************
      Call ParSpline
      ' ******************************
      '***************************************
      '*****   actualisation de la   *********
      '*****  grille des paramètres  *********
      '***************************************
      FenetreDefSpline.GridPar.Col = 1
      For i% = 1 To NbPtsSp% - 1
         FenetreDefSpline.GridPar.Row = i%
         FenetreDefSpline.GridPar.Text = Format(Aspline(i%), "0.000")
      Next i%
      FenetreDefSpline.GridPar.Col = 2
      For i% = 1 To NbPtsSp% - 1
         FenetreDefSpline.GridPar.Row = i%
         FenetreDefSpline.GridPar.Text = Format(Bspline(i%), "0.000")
      Next i%
      FenetreDefSpline.GridPar.Col = 3
      For i% = 1 To NbPtsSp% - 1
         FenetreDefSpline.GridPar.Row = i%
         FenetreDefSpline.GridPar.Text = Format(Cspline(i%), "0.000")
      Next i%
      FenetreDefSpline.GridPar.Col = 4
      For i% = 1 To NbPtsSp% - 1
         FenetreDefSpline.GridPar.Row = i%
         FenetreDefSpline.GridPar.Text = Format(Dspline(i%), "0.000")
      Next i%
      ' **************************************
      lblCalcul.Caption = "Calcul effectué !"
      Exit Sub
Traite_ErreursCalSpline:
      Select Case Err
         Case 13
            Message$ = "Erreur dans la frappe des données"
            MsgBox Message$, 48, "FenetreDefSpline.mnuCalculer"
         Case Else
            MsgBox Error$, 48, "FenetreDefSpline.mnuCalculer"
      End Select
      Exit Sub
   End If
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
   ' et écriture de leurs coordonnées Xsp() et Ysp()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbPtsSp%
   For iloc% = 1 To NbPtsSp%
      Write #1, Xsp(iloc%)
      Write #1, Ysp(iloc%)
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
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, NbPtsSp%
   Erase Xsp, Ysp
   ReDim Xsp(1 To NbPtsSp%)
   ReDim Ysp(1 To NbPtsSp%)
   ReDim Aspline(1 To NbPtsSp% - 1)
   ReDim Bspline(1 To NbPtsSp% - 1)
   ReDim Cspline(1 To NbPtsSp% - 1)
   ReDim Dspline(1 To NbPtsSp% - 1)
   For iloc% = 1 To NbPtsSp%
      Input #1, Xsp(iloc%)
      Input #1, Ysp(iloc%)
   Next iloc%
   Close #1
   '***************************************
   '***** actualisation des 2 grilles *****
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   FenetreDefSpline.txtNbPts.Text = Format(NbPtsSp%, "0")
   For i% = 1 To NbPtsSp%
      FenetreDefSpline.GridPts.Row = i%
      FenetreDefSpline.GridPts.Col = 1
      FenetreDefSpline.GridPts.Text = Format(Xsp(i%), "0.000")
      FenetreDefSpline.GridPts.Col = 2
      FenetreDefSpline.GridPts.Text = Format(Ysp(i%), "0.000")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefSpline.GridPar.Col = 1
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Aspline(i%), "0.000")
   Next i%
   FenetreDefSpline.GridPar.Col = 2
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Bspline(i%), "0.000")
   Next i%
   FenetreDefSpline.GridPar.Col = 3
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Cspline(i%), "0.000")
   Next i%
   FenetreDefSpline.GridPar.Col = 4
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(Dspline(i%), "0.000")
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
   FenetreDefSpline.Hide
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
   If NbPtsSp% < 2 Then
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "FenetreDefSpline.mnuTracer"
      Exit Sub
   ElseIf AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "FenetreDefSpline.mnuTracer"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "FenetreDefSpline.mnuTracer"
      Exit Sub
   End If
   '-----------------------------------------
   '--------   si quelque chose   -----------
   '--------    a été modifié,    -----------
   '--------   calculer d'abord   -----------
   '-----------------------------------------
   If lblCalcul.Caption = "Calcul non effectué" Then
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
   Fenetre_Prin.Caption = "Passage d'une courbe par un nuage de points (Spline)"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Détermination de la fonction Spline, union de n-1 polynômes du troisième degré, passant par n points."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.Caption = "f(X) = Union des A(I) X^3  + B(I) X^2 + C(I) X + D(I)"
   Fenetre_Val.lblNote.Caption = ""
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

Private Sub txtNbPts_Change()
   lblCalcul.Caption = "Calcul non effectué"
   NbPtsSp% = CInt(txtNbPts.Text)
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   GridPts.Rows = NbPtsSp% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   ' ********************************
   GridPts.Col = 0
   For i% = 1 To NbPtsSp%
      GridPts.Row = i%
      GridPts.Text = Format(i%, "0")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefSpline.GridPar.Rows = NbPtsSp%
   ' ********************************
   ' renumérotation 1ère colonne
   ' ********************************
   FenetreDefSpline.GridPar.Col = 0
   For i% = 1 To NbPtsSp% - 1
      FenetreDefSpline.GridPar.Row = i%
      FenetreDefSpline.GridPar.Text = Format(i%, "0")
   Next i%
   ' **************************************
End Sub


Public Sub ParSpline()
   '------------------------------------------------------
   ' *****************************************************
   ' ******  Calcul des coefficients des polynômes *******
   ' ******   du 3ième degré formant la fonction   *******
   ' ******                SPLINE                  *******
   ' *****************************************************
   ' ------------------ ReDims ---------------------------
   ReDim DXsp(1 To NbPtsSp% - 1)
   ReDim DYsp(1 To NbPtsSp% - 1)
   ReDim Aspline(1 To NbPtsSp% - 1)
   ReDim Bspline(1 To NbPtsSp% - 1)
   ReDim Cspline(1 To NbPtsSp% - 1)
   ReDim Dspline(1 To NbPtsSp% - 1)
   ReDim Espline(1 To NbPtsSp%)
   ReDim Fspline(1 To NbPtsSp%)
   ReDim Gspline(1 To NbPtsSp%)
   ReDim Hspline(1 To NbPtsSp%)
   ' -----------------------------------------------------
   ' Calcul des DX(i)=X(i+1)-X(i) et DY(i)=Y(i+1)-Y(i)
   ' -----------------------------------------------------
   For i% = 1 To NbPtsSp% - 1
      DXsp(i%) = Xsp(i% + 1) - Xsp(i%)
      DYsp(i%) = Ysp(i% + 1) - Ysp(i%)
   Next i%
   ' -----------------------------------------------------
   ' Calcul des E,F,G,H
   ' -----------------------------------------------------
   Espline(1) = 1
   Fspline(1) = 0
   Gspline(1) = 0
   Hspline(1) = 0
   For i% = 1 To NbPtsSp% - 1
      Espline(i% + 1) = -2 * Espline(i%) - Gspline(i%) * DXsp(i%) / 2
      Fspline(i% + 1) = -2 * Fspline(i%) - Hspline(i%) * DXsp(i%) / 2 + 3 * DYsp(i%) / DXsp(i%)
      Gspline(i% + 1) = -6 * Espline(i%) / DXsp(i%) - 2 * Gspline(i%)
      Hspline(i% + 1) = -6 * Fspline(i%) / DXsp(i%) - 2 * Hspline(i%) + 6 * DYsp(i%) / DXsp(i%) / DXsp(i%)
   Next i%
   AlphaSp = -Hspline(NbPtsSp%) / Gspline(NbPtsSp%)
   ' -----------------------------------------------------
   ' Calcul des A,B,C,D
   ' -----------------------------------------------------
   For i% = 1 To NbPtsSp% - 1
      Dspline(i%) = Ysp(i%)
      Cspline(i%) = Espline(i%) * AlphaSp + Fspline(i%)
      Bspline(i%) = (Gspline(i%) * AlphaSp + Hspline(i%)) / 2
      Aspline(i%) = (DYsp(i%) - Bspline(i%) * DXsp(i%) * DXsp(i%) - Cspline(i%) * DXsp(i%)) / DXsp(i%) / DXsp(i%) / DXsp(i%)
   Next i%
   ' -----------------------------------------------------
End Sub
