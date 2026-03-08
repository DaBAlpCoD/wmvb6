VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreDefPolynome 
   Caption         =   "Définition de la courbe [polynômes]"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   900
   ClientWidth     =   10005
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
   ScaleHeight     =   6615
   ScaleWidth      =   10005
   Begin MSFlexGridLib.MSFlexGrid GridPpol 
      Height          =   2295
      Left            =   120
      TabIndex        =   59
      Top             =   3960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
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
   Begin VB.PictureBox pctColorPtsChoisie 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   58
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   5640
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   4
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   5640
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   6840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   5640
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Top             =   5400
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   5400
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   6840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox PointsRelier 
      Caption         =   "relier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   720
      Value           =   2  'Grayed
      Width           =   855
   End
   Begin VB.CheckBox PointsTracer 
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
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   5880
      TabIndex        =   48
      Top             =   360
      Value           =   2  'Grayed
      Width           =   855
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
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   38
      Top             =   4680
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
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   39
      Top             =   4680
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
      Left            =   6840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   4680
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
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   42
      Top             =   4440
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
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   46
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox txtValDegP 
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
      Left            =   960
      TabIndex        =   44
      Top             =   3480
      Width           =   735
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
      Top             =   360
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
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   4440
      Width           =   495
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
      Left            =   6840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   35
      Top             =   4440
      Width           =   255
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
      Top             =   1080
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
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Value           =   2  'Grayed
      Width           =   2175
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
      Top             =   720
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
      Top             =   720
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
      Top             =   360
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
      Top             =   360
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
      Left            =   7800
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
   Begin VB.Label EtiquetteCoulPtsChoisie 
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
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   7920
      TabIndex        =   57
      Top             =   5400
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
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   5520
      TabIndex        =   50
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label EtiquettePoints 
      Caption         =   "Points :"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   5880
      TabIndex        =   47
      Top             =   120
      Width           =   735
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
   Begin VB.Label lblDegP 
      Caption         =   "Degré :"
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
      TabIndex        =   43
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblCoefP 
      Caption         =   "Coefficients :"
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
      Left            =   2040
      TabIndex        =   36
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblPol 
      Caption         =   "Polynôme "
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
      Left            =   3000
      TabIndex        =   37
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label EtiquetteGrille 
      Caption         =   "Grille :"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   120
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
      Left            =   7920
      TabIndex        =   29
      Top             =   4440
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
      Left            =   2280
      TabIndex        =   34
      Top             =   2040
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
      Height          =   615
      Left            =   6000
      TabIndex        =   33
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "P(X) = U(n-1) X^(n-1) + ... + U(1) X + U(0)"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label lblDefiniCour 
      Caption         =   "Fonction définissant la courbe :"
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
      Left            =   2880
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"DEFPOLYN.frx":0000
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
      Top             =   1560
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
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label EtiquetteAxes 
      Caption         =   "Axes :"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.Label EtiquetteCadre 
      Caption         =   "Cadre :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
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
   Begin VB.Menu mnuDefPts 
      Caption         =   "&Définir un ensemble de points"
   End
   Begin VB.Menu mnuTracer 
      Caption         =   "&Tracer..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreDefPolynome"
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
   EtiquetteFE.Caption = "     " + Format(FE, "0.0000")
   DefXmin.Text = Format(Xmin, "0.000000")
   DefXmax.Text = Format(Xmax, "0.000000")
   DefYmin.Text = Format(Ymin, "0.000000")
   DefYmax.Text = Format(Ymax, "0.000000")
   '***************************************
   '***** actualisation de la grille ******
   '***************************************
   FenetreDefPolynome.gridPpol.Rows = DegPpol% + 2
   FenetreDefPolynome.gridPpol.Cols = 2
   ' ********************************
   ' placement des valeurs
   ' ********************************
   FenetreDefPolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetreDefPolynome.gridPpol.Row = i% + 1
      FenetreDefPolynome.gridPpol.Text = Format(Ppol(i%), "0.000")
   Next i%
   ' *****************************************
   ' ****  placement du degré du polynôme ****
   ' *****************************************
   FenetreDefPolynome.txtValDegP.Text = Format(DegPpol%, "0")
   ' ATTENTION ! ceci provoque un ReDim
   ' qui efface les Ppol(i%)
   ' *****************************************
End Sub

Private Sub Form_Deactivate()
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetreDefPolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetreDefPolynome.gridPpol.Row = i% + 1
      If FenetreDefPolynome.gridPpol.Text = "" Then
         FenetreDefPolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetreDefPolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
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
   '**************************************
   '***** mise en place de la grille *****
   '**************************************
   FenetreDefPolynome.gridPpol.Rows = DegPpol% + 2
   FenetreDefPolynome.gridPpol.Cols = 2
   FenetreDefPolynome.gridPpol.FixedAlignment(0) = 2
   FenetreDefPolynome.gridPpol.FixedAlignment(1) = 2
   FenetreDefPolynome.gridPpol.ColWidth(1) = 1200
   FenetreDefPolynome.gridPpol.Row = 0
   FenetreDefPolynome.gridPpol.Col = 0
   FenetreDefPolynome.gridPpol.Text = "p"
   FenetreDefPolynome.gridPpol.Col = 1
   FenetreDefPolynome.gridPpol.Text = "Up"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefPolynome.gridPpol.Col = 0
   For i% = 0 To DegPpol%
      FenetreDefPolynome.gridPpol.Row = i% + 1
      FenetreDefPolynome.gridPpol.Text = Format(i%, "0")
   Next i%
   '*********************************
End Sub

Private Sub gridPpol_KeyPress(KeyAscii As Integer)
   ElePText$ = gridPpol.Text
   Select Case KeyAscii
   Case 32 To 168
      ElePCar$ = Chr(KeyAscii)
      ElePText$ = ElePText$ & ElePCar$
      gridPpol.Text = ElePText$
   Case 8
      If Len(gridPpol.Text) > 0 Then
         ElePText$ = Left$(ElePText$, Len(ElePText$) - 1)
         gridPpol.Text = ElePText$
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


Private Sub mnuDefPts_Click()
   FenetreDefPts.Show
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
   ' et écriture de leurs coordonnées Xpt() et Ypt()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbPts%
   For iloc% = 1 To NbPts%
      Write #1, Xpt(iloc%)
      Write #1, Ypt(iloc%)
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
   Input #1, NbPts%
   ReDim Xpt(1 To NbPts%)
   ReDim Ypt(1 To NbPts%)
   For iloc% = 1 To NbPts%
    Input #1, Xpt(iloc%)
    Input #1, Ypt(iloc%)
   Next iloc%
   Close #1
   '-----------------------------------------------------
   ' activation de certaines cases
   '-------------------------------
   EtiquettePoints.ForeColor = &H80000012     'Texte activé
   PointsTracer.ForeColor = &H80000012        'Texte activé
   PointsTracer.Value = 1                         'Checked
   PointsRelier.ForeColor = &H80000012        'Texte activé
   PointsRelier.Value = 0                         'Unchecked
   EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
   EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
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
   FenetreDefPolynome.Hide
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
   If DegPpol% < 0 Then
      MsgBox "le degré de P doit être supérieur ou égal à 0 !", 48, "FenetreDefPolynome.mnuTracer"
      Exit Sub
   ElseIf AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "FenetreDefPolynome.mnuTracer"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "FenetreDefPolynome.mnuTracer"
      Exit Sub
   End If
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetreDefPolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetreDefPolynome.gridPpol.Row = i% + 1
      If FenetreDefPolynome.gridPpol.Text = "" Then
         FenetreDefPolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetreDefPolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
   '-----------------------------------------------------
   '--------        Ouverture            ----------------
   '--------  des fenêtres pour le tracé ----------------
   '-----------------------------------------------------
   Fenetre_Prin.Show
   Fenetre_Val.Show
   Dessin.Show
   '-----------------------------------------------------
   '--------- Affichage de la formule -------------------
   '---------  définissant la courbe  -------------------
   '-----------------------------------------------------
   Fenetre_Prin.Caption = "Polynôme"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Tracé d'une courbe définie par un polynôme."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   TextVar2$ = Format(Ppol(DegPpol%), "0.00") & " X^" & Format(DegPpol%, "0")
   If DegPpol% > 1 Then
      For i% = DegPpol% - 1 To 1 Step -1
         TextVar2$ = TextVar2$ & " + " & Format(Ppol(i%), "0.00") & " X^" & Format(i%, "0")
      Next i%
   End If
   TextVar2$ = TextVar2$ & " + " & Format(Ppol(0), "0.00")
   Fenetre_Val.lblFormule.Caption = "F(X) = " & TextVar2$
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
   pctColorPtsChoisie.BackColor = pctColorPts(Index).BackColor
   CouleurPts& = pctColorPtsChoisie.BackColor
End Sub

Private Sub PointsRelier_Click()
    Rep% = PointsRelier.Value
End Sub

Private Sub PointsTracer_Click()
    Trap% = PointsTracer.Value
End Sub

Private Sub txtValDegP_Change()
   On Error Resume Next
   If txtValDegP = "" Then
      DegPpol% = 0
   Else
      DegPpol% = CInt(txtValDegP.Text)
   End If
   If DegPpol% < 0 Then
      Beep
      MsgBox "le degré de P doit être supérieur ou égal à 0 !", 48, "POLYNOME"
      DegPpol% = 0
      txtValDegP.Text = "0"
   End If
   gridPpol.Rows = DegPpol% + 2
   '***************************************
   ReDim Ppol(0 To DegPpol%)
   '***************************************
   ' renumérotation 1ère colonne
   ' ********************************
   FenetreDefPolynome.gridPpol.Col = 0
   For i% = 0 To DegPpol%
      FenetreDefPolynome.gridPpol.Row = i% + 1
      FenetreDefPolynome.gridPpol.Text = Format(i%, "0")
   Next i%
   ' ********************************
   On Error GoTo 0
End Sub

