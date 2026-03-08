VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreDefInPoLag 
   Caption         =   "Définition de la courbe [polynômes d'interpolation de Lagrange]"
   ClientHeight    =   6870
   ClientLeft      =   270
   ClientTop       =   870
   ClientWidth     =   10305
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
   ScaleHeight     =   6870
   ScaleWidth      =   10305
   Begin MSFlexGridLib.MSFlexGrid GridPts 
      Height          =   1215
      Left            =   4680
      TabIndex        =   64
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
   Begin MSFlexGridLib.MSFlexGrid GridPar 
      Height          =   1215
      Left            =   240
      TabIndex        =   63
      Top             =   4080
      Width           =   4215
      _ExtentX        =   7435
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
      TabIndex        =   62
      Top             =   600
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
      TabIndex        =   61
      Top             =   240
      Value           =   2  'Grayed
      Width           =   855
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
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   6480
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
      Left            =   4680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   58
      Top             =   6480
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
      Left            =   4440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   57
      Top             =   6480
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
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   6240
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
      Left            =   4680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   6240
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
      Left            =   4440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   6240
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
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Top             =   6480
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
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   6480
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
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Top             =   6480
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
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   6240
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
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   6240
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
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   48
      Top             =   6240
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
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   6240
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
      Left            =   2640
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   6240
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
      Width           =   2295
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
      Left            =   8040
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
   Begin VB.Label EtiquettePoints 
      Caption         =   "Points :"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   5880
      TabIndex        =   60
      Top             =   0
      Width           =   735
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
      Caption         =   $"DEFPOLAG.frx":0000
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
      Left            =   5400
      TabIndex        =   38
      Top             =   6240
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
      Left            =   3720
      TabIndex        =   37
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblPar 
      Caption         =   "Paramètres entrant dans la définition de la courbe"
      Height          =   255
      Left            =   240
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
      Left            =   1800
      TabIndex        =   29
      Top             =   6240
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
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "f(X) = U(n-1) X^(n-1) + ... + U(1) X + U(0)"
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
      Width           =   3855
   End
   Begin VB.Label lblDefiniCour 
      Caption         =   "Fonction définissant la courbe :"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   2895
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
      Left            =   2760
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
Attribute VB_Name = "FenetreDefInPoLag"
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
   DefYmax.Text = Format(Xmax, "0.000000")
   '***************************************
   '***** actualisation des 2 grilles *****
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   FenetreDefInPoLag.txtNbPts.Text = Format(NbPtsI%, "0")
   For i% = 1 To NbPtsI%
      FenetreDefInPoLag.gridPts.Row = i%
      FenetreDefInPoLag.gridPts.Col = 1
      FenetreDefInPoLag.gridPts.Text = Format(Xpl(i%), "0.0000")
      FenetreDefInPoLag.gridPts.Col = 2
      FenetreDefInPoLag.gridPts.Text = Format(Ypl(i%), "0.0000")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefInPoLag.gridPar.Col = 1
   For i% = 0 To DegPoLag%
      FenetreDefInPoLag.gridPar.Row = i% + 1
      FenetreDefInPoLag.gridPar.Text = Format(UparI(i%), "0.000000")
   Next i%
   ' **************************************
End Sub

Private Sub Form_Deactivate()
   On Error GoTo Traite_ErreursFoDePoLag
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Xpl, Ypl, UparI
   ReDim Xpl(1 To NbPtsI%), Ypl(1 To NbPtsI%)
   ReDim UparI(0 To DegPoLag%)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsI%
      FenetreDefInPoLag.gridPts.Row = i%
      FenetreDefInPoLag.gridPts.Col = 1
      If FenetreDefInPoLag.gridPts.Text = "" Then
         FenetreDefInPoLag.gridPts.Text = "0"
         Xpl(i%) = 0
      Else
         Xpl(i%) = CSng(FenetreDefInPoLag.gridPts.Text)
      End If
      FenetreDefInPoLag.gridPts.Col = 2
      If FenetreDefInPoLag.gridPts.Text = "" Then
         FenetreDefInPoLag.gridPts.Text = "0"
         Ypl(i%) = 0
      Else
         Ypl(i%) = CSng(FenetreDefInPoLag.gridPts.Text)
      End If
   Next i%
   ' ********************************************
   ' affectation de leurs valeurs aux paramètres
   ' ********************************************
   FenetreDefInPoLag.gridPar.Col = 1
   For i% = 0 To DegPoLag%
      FenetreDefInPoLag.gridPar.Row = i% + 1
      If FenetreDefInPoLag.gridPar.Text = "" Then
         FenetreDefInPoLag.gridPar.Text = "0"
         UparI(i%) = 0
      Else
         UparI(i%) = CSng(FenetreDefInPoLag.gridPar.Text)
      End If
   Next i%
   ' **************************************
   Exit Sub
Traite_ErreursFoDePoLag:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48, "FenetreDefInPoLag.Deactivate"
      Case Else
         MsgBox Error$, 48, "FenetreDefInPoLag.Deactivate"
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
   '----------------    pour InPoLag    -----------------
   '-----------------------------------------------------
   lblCalcul.Caption = "Calcul non effectué"
   NbPtsI% = 7    ' nombre de points
   DegPoLag% = NbPtsI% - 1 ' Degré du polynôme de Lagrange
   ReDim Xpl(1 To NbPtsI%), Ypl(1 To NbPtsI%)
   ReDim UparI(0 To DegPoLag%)
   '************ valeurs des paramètres **************
   For i% = 0 To DegPoLag%
      UparI(i%) = 0
   Next i%
   '**** points sur y = .02x^6-.02x^5-.04x^4+.37x^3+2.18x^2-1.48x-1.56 ****
   Xpl(1) = -3.5
   Ypl(1) = 3
   Xpl(2) = -2.8
   Ypl(2) = 0.5
   Xpl(3) = -1.2
   Ypl(3) = 2
   Xpl(4) = 0.7
   Ypl(4) = -1.5
   Xpl(5) = 1.8
   Ypl(5) = 1
   Xpl(6) = 2.5
   Ypl(6) = 1
   Xpl(7) = 3.2
   Ypl(7) = -0.5
   '***************************************
   '***** mise en place des 2 grilles *****
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   FenetreDefInPoLag.txtNbPts.Text = Format(NbPtsI%, "0")
   FenetreDefInPoLag.gridPts.Cols = 3
   FenetreDefInPoLag.gridPts.Rows = NbPtsI% + 1
   FenetreDefInPoLag.gridPts.FixedAlignment(0) = 2
   FenetreDefInPoLag.gridPts.FixedAlignment(1) = 2
   FenetreDefInPoLag.gridPts.FixedAlignment(2) = 2
   FenetreDefInPoLag.gridPts.ColWidth(1) = 1500
   FenetreDefInPoLag.gridPts.ColWidth(2) = 1500
   FenetreDefInPoLag.gridPts.Row = 0
   FenetreDefInPoLag.gridPts.Col = 1
   FenetreDefInPoLag.gridPts.Text = "X"
   FenetreDefInPoLag.gridPts.Col = 2
   FenetreDefInPoLag.gridPts.Text = "Y"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefInPoLag.gridPts.Col = 0
   For i% = 1 To NbPtsI%
      FenetreDefInPoLag.gridPts.Row = i%
      FenetreDefInPoLag.gridPts.Text = Format(i%, "0")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefInPoLag.gridPar.Cols = 2
   FenetreDefInPoLag.gridPar.Rows = DegPoLag% + 2
   FenetreDefInPoLag.gridPar.FixedAlignment(0) = 2
   FenetreDefInPoLag.gridPar.FixedAlignment(1) = 2
   FenetreDefInPoLag.gridPar.ColWidth(1) = 2500
   FenetreDefInPoLag.gridPar.Row = 0
   FenetreDefInPoLag.gridPar.Col = 1
   FenetreDefInPoLag.gridPar.Text = "U"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefInPoLag.gridPar.Col = 0
   For i% = 0 To DegPoLag%
      FenetreDefInPoLag.gridPar.Row = i% + 1
      FenetreDefInPoLag.gridPar.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub

Private Sub gridPts_KeyPress(KeyAscii As Integer)
   lblCalcul.Caption = "Calcul non effectué"
   PtsText$ = gridPts.Text
   Select Case KeyAscii
   Case 32 To 168
      PtsCar$ = Chr(KeyAscii)
      PtsText$ = PtsText$ & PtsCar$
      gridPts.Text = PtsText$
   Case 8
      If Len(gridPts.Text) > 0 Then
         PtsText$ = Left$(PtsText$, Len(PtsText$) - 1)
         gridPts.Text = PtsText$
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
      On Error GoTo Traite_ErreursCalPoLag
      ' ********************************************
      '               ReDims
      ' ********************************************
      Erase Xpl, Ypl, UparI
      ReDim Xpl(1 To NbPtsI%), Ypl(1 To NbPtsI%)
      ReDim UparI(0 To DegPoLag%)
      ' ********************************************
      ' affectation de leurs valeurs aux coordonnées
      ' ********************************************
      For i% = 1 To NbPtsI%
         FenetreDefInPoLag.gridPts.Row = i%
         FenetreDefInPoLag.gridPts.Col = 1
         If FenetreDefInPoLag.gridPts.Text = "" Then
            FenetreDefInPoLag.gridPts.Text = "0"
            Xpl(i%) = 0
         Else
            Xpl(i%) = CSng(FenetreDefInPoLag.gridPts.Text)
         End If
         FenetreDefInPoLag.gridPts.Col = 2
         If FenetreDefInPoLag.gridPts.Text = "" Then
            FenetreDefInPoLag.gridPts.Text = "0"
            Ypl(i%) = 0
         Else
            Ypl(i%) = CSng(FenetreDefInPoLag.gridPts.Text)
         End If
      Next i%
      ' ********************************************
      ' Calcul des paramètres Upar
      ' par l'algorithme d' AITKEN
      ' ********************************************
      Call Aitken
      '***************************************
      '*****  actualisation de la   **********
      '***** grille des paramètres  **********
      '***************************************
      FenetreDefInPoLag.gridPar.Col = 1
      For i% = 0 To DegPoLag%
         FenetreDefInPoLag.gridPar.Row = i% + 1
         FenetreDefInPoLag.gridPar.Text = Format(UparI(i%), "0.000000")
      Next i%
      ' **************************************
      lblCalcul.Caption = "Calcul effectué !"
      Exit Sub
Traite_ErreursCalPoLag:
      Select Case Err
         Case 13
            Message$ = "Erreur dans la frappe des données"
            MsgBox Message$, 48, "FenetreDefInPoLag.mnuCalculer"
         Case Else
            MsgBox Error$, 48, "FenetreDefInPoLag.mnuCalculer"
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
   ' et écriture de leurs coordonnées Xpt() et Ypt()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbPtsI%
   For iloc% = 1 To NbPtsI%
      Write #1, Xpl(iloc%)
      Write #1, Ypl(iloc%)
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
   Input #1, NbPtsI%
   DegPoLag% = NbPtsI% - 1
   ReDim Xpl(1 To NbPtsI%)
   ReDim Ypl(1 To NbPtsI%)
   ReDim UparI(0 To NbPtsI% - 1)
   For iloc% = 1 To NbPtsI%
      Input #1, Xpl(iloc%)
      Input #1, Ypl(iloc%)
   Next iloc%
   Close #1
   '***************************************
   '***** actualisation des 2 grilles *****
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   FenetreDefInPoLag.txtNbPts.Text = Format(NbPtsI%, "0")
   For i% = 1 To NbPtsI%
      FenetreDefInPoLag.gridPts.Row = i%
      FenetreDefInPoLag.gridPts.Col = 1
      FenetreDefInPoLag.gridPts.Text = Format(Xpl(i%), "0.0000")
      FenetreDefInPoLag.gridPts.Col = 2
      FenetreDefInPoLag.gridPts.Text = Format(Ypl(i%), "0.0000")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefInPoLag.gridPar.Col = 1
   For i% = 0 To DegPoLag%
      FenetreDefInPoLag.gridPar.Row = i% + 1
      FenetreDefInPoLag.gridPar.Text = Format(UparI(i%), "0.000000")
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
   FenetreDefInPoLag.Hide
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
   If NbPtsI% < 2 Then
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "FenetreDefInPoLag.mnuTracer"
      Exit Sub
   ElseIf AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "FenetreDefInPoLag.mnuTracer"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "FenetreDefInPoLag.mnuTracer"
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
   '--------        Ouverture           ----------------
   '--------  des fenêtres pour le tracé ----------------
   '-----------------------------------------------------
   Fenetre_Prin.Show
   Fenetre_Val.Show
   Dessin.Show
   '--------------------------------------------------------------
   '--------- Affichage de la formule ----------------------------
   '---------     définissant la courbe        -------------------
   '--------------------------------------------------------------
   Fenetre_Prin.Caption = "Passage d'une courbe par un nuage de points (polynômes de Lagrange)"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Détermination du polynôme de Lagrange de degré n-1 passant par n points." & Chr$(13) & "On utilise l'algorithme d'Aitken."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   TextVar2$ = Format(UparI(DegPoLag%), "0.00") & " X^" & Format(DegPoLag%, "0")
   If DegPoLag% > 1 Then
      For i% = DegPoLag% - 1 To 1 Step -1
         TextVar2$ = TextVar2$ & " + " & Format(UparI(i%), "0.00") & " X^" & Format(i%, "0")
      Next i%
   End If
   TextVar2$ = TextVar2$ & " + " & Format(UparI(0), "0.00")
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
   pctColorPoints.BackColor = pctColorPts(Index).BackColor
   CouleurPts& = pctColorPoints.BackColor
End Sub

Private Sub txtNbPts_Change()
   lblCalcul.Caption = "Calcul non effectué"
   NbPtsI% = CInt(txtNbPts.Text)
   DegPoLag% = NbPtsI% - 1
   '***************************************
   '***** 1) grille des points ************
   '***************************************
   gridPts.Rows = NbPtsI% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   ' ********************************
   gridPts.Col = 0
   For i% = 1 To NbPtsI%
      gridPts.Row = i%
      gridPts.Text = Format(i%, "0")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefInPoLag.gridPar.Rows = DegPoLag% + 2
   ' ********************************
   ' renumérotation 1ère colonne
   ' ********************************
   FenetreDefInPoLag.gridPar.Col = 0
   For i% = 0 To DegPoLag%
      FenetreDefInPoLag.gridPar.Row = i% + 1
      FenetreDefInPoLag.gridPar.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub


Public Sub CalculerPoLag()
   Call Aitken
End Sub

Public Sub Aitken()
   '------------------------------------------------------
   ' *****************************************************
   ' ****************  Algorithme d' AITKEN  *************
   ' *****************************************************
   ' ------------------ ReDims ---------------------------
   ReDim Fpolag(0 To DegPoLag%)
   ReDim Gpolag(0 To DegPoLag%)
   ReDim Hpolag(0 To DegPoLag%)
   ' -----------------------------------------------------
   Fpolag(0) = Ypl(1)
   For i% = 0 To DegPoLag%
      Gpolag(i%) = Ypl(i% + 1)
   Next i%
   For i% = 1 To DegPoLag%
      For j% = 0 To DegPoLag% - i%
         Hpolag(j%) = (Gpolag(j% + 1) - Gpolag(j%)) / (Xpl(j% + i% + 1) - Xpl(j% + 1))
         Gpolag(j%) = Hpolag(j%)
      Next j%
      Fpolag(i%) = Hpolag(0)
   Next i%
   ' -----------------------------------------------------
   ' *****************************************************
   ' *************  Calcul des coefficients  *************
   ' *************  du polynôme de Lagrange  *************
   ' *****************************************************
   UparI(0) = Fpolag(DegPoLag%)
   For i% = 1 To DegPoLag%
      UparI(i%) = 0
   Next i%
   For i% = 1 To DegPoLag%
      For j% = i% To 1 Step -1
         UparI(j%) = UparI(j% - 1) - Xpl(DegPoLag% - i% + 1) * UparI(j%)
      Next j%
      UparI(0) = Fpolag(DegPoLag% - i%) - Xpl(DegPoLag% - i% + 1) * UparI(0)
   Next i%
   ' -----------------------------------------------------
End Sub

Public Sub MemPoLag()
On Error GoTo Traite_ErreursMemPoLag
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Xpl, Ypl, UparI
   ReDim Xpl(1 To NbPtsI%), Ypl(1 To NbPtsI%)
   ReDim UparI(0 To DegPoLag%)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsI%
      FenetreDefInPoLag.gridPts.Row = i%
      FenetreDefInPoLag.gridPts.Col = 1
      If FenetreDefInPoLag.gridPts.Text = "" Then
         FenetreDefInPoLag.gridPts.Text = "0"
         Xpl(i%) = 0
      Else
         Xpl(i%) = CSng(FenetreDefInPoLag.gridPts.Text)
      End If
      FenetreDefInPoLag.gridPts.Col = 2
      If FenetreDefInPoLag.gridPts.Text = "" Then
         FenetreDefInPoLag.gridPts.Text = "0"
         Ypl(i%) = 0
      Else
         Ypl(i%) = CSng(FenetreDefInPoLag.gridPts.Text)
      End If
   Next i%
   ' ********************************************
   ' affectation de leurs valeurs aux paramètres
   ' ********************************************
   FenetreDefInPoLag.gridPar.Col = 1
   For i% = 0 To DegPoLag%
      FenetreDefInPoLag.gridPar.Row = i% + 1
      If FenetreDefInPoLag.gridPar.Text = "" Then
         FenetreDefInPoLag.gridPar.Text = "0"
         UparI(i%) = 0
      Else
         UparI(i%) = CSng(FenetreDefInPoLag.gridPar.Text)
      End If
   Next i%
   Exit Sub
Traite_ErreursMemPoLag:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub
