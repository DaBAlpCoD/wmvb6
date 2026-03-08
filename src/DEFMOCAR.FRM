VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreDefMoinCar 
   Caption         =   "Définition de la courbe [affinement par moindres carrés]"
   ClientHeight    =   7215
   ClientLeft      =   255
   ClientTop       =   615
   ClientWidth     =   9435
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
   ScaleHeight     =   7215
   ScaleWidth      =   9435
   Begin MSFlexGridLib.MSFlexGrid GridPts 
      Height          =   1215
      Left            =   4680
      TabIndex        =   47
      Top             =   4560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   3
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
      TabIndex        =   48
      Top             =   4560
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
      Height          =   255
      Left            =   4680
      TabIndex        =   65
      Top             =   960
      Value           =   1  'Checked
      Width           =   975
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
      Left            =   4440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   63
      Top             =   6120
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
      Left            =   4200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   62
      Top             =   6120
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
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   61
      Top             =   6120
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
      Left            =   4440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   60
      Top             =   5880
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
      Left            =   4200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   59
      Top             =   5880
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
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   58
      Top             =   5880
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
      TabIndex        =   57
      Top             =   6120
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
      TabIndex        =   56
      Top             =   6120
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
      TabIndex        =   55
      Top             =   6120
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
      TabIndex        =   54
      Top             =   5880
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
      TabIndex        =   53
      Top             =   5880
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
      TabIndex        =   52
      Top             =   5880
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
      TabIndex        =   46
      Top             =   4200
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
      Left            =   5640
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtNbPar 
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
      Left            =   1080
      TabIndex        =   38
      Top             =   4200
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
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox DefVarf 
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
      Left            =   840
      TabIndex        =   27
      Text            =   "U(1)*EXP(U(2)*X*X)"
      Top             =   3480
      Width           =   8415
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
      Width           =   2055
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
      Left            =   6120
      TabIndex        =   31
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label EtiquettePoints 
      Caption         =   "Points :"
      Height          =   255
      Left            =   4680
      TabIndex        =   64
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblValEr 
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
      Left            =   7200
      TabIndex        =   51
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblErreur 
      Caption         =   "Valeur de l'erreur :"
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
      Left            =   6360
      TabIndex        =   50
      Top             =   5880
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   3960
      Y2              =   5760
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
      TabIndex        =   49
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
      TabIndex        =   45
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblPts 
      Caption         =   "Points par où on cherche à faire passer la courbe"
      Height          =   255
      Left            =   4680
      TabIndex        =   44
      Top             =   3960
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
      TabIndex        =   43
      Top             =   4200
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
      TabIndex        =   36
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblNbPar 
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
      Left            =   120
      TabIndex        =   42
      Top             =   4200
      Width           =   855
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
      Left            =   4920
      TabIndex        =   40
      Top             =   5880
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
      Left            =   3240
      TabIndex        =   39
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label lblPar 
      Caption         =   "Paramètres entrant dans la définition de la courbe"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label EtiquetteGrille 
      Caption         =   "Grille :"
      Height          =   255
      Left            =   4680
      TabIndex        =   29
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
      TabIndex        =   30
      Top             =   5880
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "F(X) ="
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
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   735
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
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"DEFMOCAR.frx":0000
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
      TabIndex        =   28
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
Attribute VB_Name = "FenetreDefMoinCar"
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



Private Sub Check1_Click()

End Sub

Private Sub DefVarf_Change()
   TextVar2$ = DefVarf.Text
   FenetreDefMoinCar.lblValEr.Caption = ""
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
    DefVarf.Text = TextVar2$
   '***************************************
   '***** actualisation de la formule *****
   '***************************************
   FenetreDefMoinCar.lblValEr.Caption = ""
   DefVarf.Text = TextVar2$
   '***************************************
   '***** actualisation des 2 grilles *****
   '***** 1) grille des points ************
   '***************************************
   FenetreDefMoinCar.txtNbPts.Text = Format(NbPtsMoinCar%, "0")
   For i% = 1 To NbPtsMoinCar%
      FenetreDefMoinCar.GridPts.Row = i%
      FenetreDefMoinCar.GridPts.Col = 1
      FenetreDefMoinCar.GridPts.Text = Format(Xmc(i%), "0.0000")
      FenetreDefMoinCar.GridPts.Col = 2
      FenetreDefMoinCar.GridPts.Text = Format(Ymc(i%), "0.0000")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefMoinCar.txtNbPar.Text = Format(NbPar%, "0")
   FenetreDefMoinCar.GridPar.Col = 1
   For i% = 1 To NbPar%
      FenetreDefMoinCar.GridPar.Row = i%
      FenetreDefMoinCar.GridPar.Text = Format(Upar(i%), "0.000000")
   Next i%
   ' **************************************
End Sub

Private Sub Form_Deactivate()
   Call Affecte
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
   '----------------    pour MoinCar    -----------------
   '-----------------------------------------------------
   TextVar2$ = "U(1)*EXP(U(2)*X*X)"
   NbPtsMoinCar% = 20    ' nombre de points
   NbPar% = 2  ' nombre de paramètres
   ReDim Xmc(1 To NbPtsMoinCar%), Ymc(1 To NbPtsMoinCar%)
   ReDim Upar(1 To NbPar%)
   ReDim UApar(1 To NbPar%)
   '************ valeurs des paramètres ************
   Upar(1) = 3
   UApar(1) = 3
   Upar(2) = -0.7
   UApar(2) = -0.7
   '************ points sur y = 4 exp(-0.5 x²)
   Xmc(1) = -3.26
   Ymc(1) = 0.022
   Xmc(2) = -2.48
   Ymc(2) = 0.18
   Xmc(3) = -1.91
   Ymc(3) = 0.64
   Xmc(4) = -1.64
   Ymc(4) = 1.02
   Xmc(5) = -1.31
   Ymc(5) = 1.66
   Xmc(6) = -1.2
   Ymc(6) = 1.98
   Xmc(7) = -0.97
   Ymc(7) = 2.48
   Xmc(8) = -0.84
   Ymc(8) = 2.8
   Xmc(9) = -0.53
   Ymc(9) = 3.46
   Xmc(10) = -0.26
   Ymc(10) = 3.88
   Xmc(11) = 0.11
   Ymc(11) = 3.97
   Xmc(12) = 0.35
   Ymc(12) = 3.75
   Xmc(13) = 0.71
   Ymc(13) = 3.13
   Xmc(14) = 1
   Ymc(14) = 2.42
   Xmc(15) = 1.22
   Ymc(15) = 1.91
   Xmc(16) = 1.6
   Ymc(16) = 1.13
   Xmc(17) = 1.82
   Ymc(17) = 0.78
   Xmc(18) = 2.22
   Ymc(18) = 0.35
   Xmc(19) = 2.7
   Ymc(19) = 0.11
   Xmc(20) = 3.5
   Ymc(20) = 0.02
   '-----------------------------------------------------
   '***************************************
   '***** mise en place des 2 grilles *****
   '***** 1) grille des points ************
   ' **************************************
   FenetreDefMoinCar.txtNbPts.Text = Format(NbPtsMoinCar%, "0")
   FenetreDefMoinCar.GridPts.Rows = NbPtsMoinCar% + 1
   FenetreDefMoinCar.GridPts.Cols = 3
   FenetreDefMoinCar.GridPts.FixedAlignment(0) = 2
   FenetreDefMoinCar.GridPts.FixedAlignment(1) = 2
   FenetreDefMoinCar.GridPts.FixedAlignment(2) = 2
   FenetreDefMoinCar.GridPts.ColWidth(1) = 1500
   FenetreDefMoinCar.GridPts.ColWidth(2) = 1500
   FenetreDefMoinCar.GridPts.Row = 0
   FenetreDefMoinCar.GridPts.Col = 1
   FenetreDefMoinCar.GridPts.Text = "X"
   FenetreDefMoinCar.GridPts.Col = 2
   FenetreDefMoinCar.GridPts.Text = "Y"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefMoinCar.GridPts.Col = 0
   For i% = 1 To NbPtsMoinCar%
      FenetreDefMoinCar.GridPts.Row = i%
      FenetreDefMoinCar.GridPts.Text = Format(i%, "0")
   Next i%
   ' **************************************
   ' ***** 2) grille des paramètres  ******
   ' **************************************
   FenetreDefMoinCar.txtNbPar.Text = Format(NbPar%, "0")
   FenetreDefMoinCar.GridPar.Rows = NbPar% + 1
   FenetreDefMoinCar.GridPar.Cols = 2
   FenetreDefMoinCar.GridPar.FixedAlignment(0) = 2
   FenetreDefMoinCar.GridPar.FixedAlignment(1) = 2
   FenetreDefMoinCar.GridPar.ColWidth(1) = 2500
   FenetreDefMoinCar.GridPar.Row = 0
   FenetreDefMoinCar.GridPar.Col = 1
   FenetreDefMoinCar.GridPar.Text = "U"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreDefMoinCar.GridPar.Col = 0
   For i% = 1 To NbPar%
      FenetreDefMoinCar.GridPar.Row = i%
      FenetreDefMoinCar.GridPar.Text = Format(i%, "0")
   Next i%
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
   FenetreDefMoinCar.lblValEr.Caption = ""
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

Public Sub MoinCar()
   '----------------------------------------------------------
   ' Affinement par moindres carrés :
   ' --------------------------------
   ' Equations normales :  Mmat(i,j) . Uvec(j) = Vvec(i)
   ' On cherche les Uvec(j);
   ' Mmat est symétrique définie positive et peut donc
   ' être inversée par la méthode de CHOLESKI.
   ' ---------------------------------------------------------
   On Error GoTo Traite_ErreursMoCar
   '--------------- Calcul des Mmat(i,j) et Vvec(j) ----------
   ' Initialisations
   OrdreMat% = NbPar%
   ' ********************************************
   '                ReDims
   ' ********************************************
   ReDim Mmat(1 To NbPar%, 1 To NbPar%)
   ReDim Uvec(1 To NbPar%)
   ReDim Vvec(1 To NbPar%)
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Xmc, Ymc, Upar, UApar
   ReDim Xmc(1 To NbPtsMoinCar%), Ymc(1 To NbPtsMoinCar%)
   ReDim Upar(1 To NbPar%)
   ReDim UApar(1 To NbPar%)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsMoinCar%
      FenetreDefMoinCar.GridPts.Row = i%
      FenetreDefMoinCar.GridPts.Col = 1
      Xmc(i%) = CSng(FenetreDefMoinCar.GridPts.Text)
      FenetreDefMoinCar.GridPts.Col = 2
      Ymc(i%) = CSng(FenetreDefMoinCar.GridPts.Text)
   Next i%
   ' ********************************************
   ' affectation de leurs valeurs aux paramètres
   ' ********************************************
   FenetreDefMoinCar.GridPar.Col = 1
   For i% = 1 To NbPar%
      FenetreDefMoinCar.GridPar.Row = i%
      Upar(i%) = CSng(FenetreDefMoinCar.GridPar.Text)
      UApar(i%) = CSng(FenetreDefMoinCar.GridPar.Text)
   Next i%
   ' ********************************************
   eps = 0.001
   ' ********************************************
   ' initialisation
   For i% = 1 To NbPar%
      Vvec(i%) = 0
      For j% = 1 To NbPar%
         Mmat(i%, j%) = 0
      Next j%
   Next i%
   ' ********************************************
   For i% = 1 To NbPar%
      For k% = 1 To NbPtsMoinCar%
         X = Xmc(k%)
         ' ++++++++++++++++++++++++++++++++++++++++
         ' ++++++++++  dyi = df(Xk)/dUi  ++++++++++
         ' ++++++++++++++++++++++++++++++++++++++++
         Upar(i%) = UApar(i%)
         vvar$(1) = LTrim$(Format(X, "0.000000"))
         Call Parametre(ForVar2$, ParVar2$)
         Call variable(ParVar2$, ValVar2$)
         Call Traitement(ValVar2$, SorVar2$)
         Ymc0 = CSng(SorVar2$)
         Upar(i%) = UApar(i%) + eps
         vvar$(1) = LTrim$(Format(X))
         Call Parametre(ForVar2$, ParVar2$)
         Call variable(ParVar2$, ValVar2$)
         Call Traitement(ValVar2$, SorVar2$)
         Ymc1 = CSng(SorVar2$)
         dyi = (Ymc1 - Ymc0) / eps
         ' ++++++++++++++++++++++++++++++++++++++++
         ' Vvec(i)=somme des [Yk-f(Xk)](df(Xk)/dUi)
         ' ++++++++++++++++++++++++++++++++++++++++
         Vvec(i%) = Vvec(i%) + (Ymc(k%) - Ymc0) * dyi
         ' ++++++++++++++++++++++++++++++++++++++++
         For j% = i% To NbPar%
            ' ++++++++++++++++++++++++++++++++++++++++
            ' ++++++++++  dyj = df(Xk)/dUj  ++++++++++
            ' ++++++++++++++++++++++++++++++++++++++++
            Upar(j%) = UApar(j%)
            vvar$(1) = LTrim$(Format(X, "0.000000"))
            Call Parametre(ForVar2$, ParVar2$)
            Call variable(ParVar2$, ValVar2$)
            Call Traitement(ValVar2$, SorVar2$)
            Ymc0 = CSng(SorVar2$)
            Upar(j%) = UApar(j%) + eps
            vvar$(1) = LTrim$(Format(X, "0.000000"))
            Call Parametre(ForVar2$, ParVar2$)
            Call variable(ParVar2$, ValVar2$)
            Call Traitement(ValVar2$, SorVar2$)
            Ymc1 = CSng(SorVar2$)
            dyj = (Ymc1 - Ymc0) / eps
            ' ++++++++++++++++++++++++++++++++++++++++++++
            ' Mmat(i,j)=somme des (df(Xk)/dUi)(df(Xk)/dUj)
            ' ++++++++++++++++++++++++++++++++++++++++++++
            Mmat(i%, j%) = Mmat(i%, j%) + dyi * dyj
            ' ++++++++++++++++++++++++++++++++++++++++++++
         Next j%
      Next k%
   Next i%
   '--------------- Calcul des Uvec(i) ----------
   Call InvMatCholeski(OrdreMat%, DePo, DetMmat, Mmat(), Wmat())
   If DePo = False Then
      Message$ = "Matrice non définie positive"
      MsgBox Message$, 48
      Exit Sub
   End If
   ' Uvec(i) = Wmat(i,j) . Vvec(j)
   For i% = 1 To NbPar%
      elem = 0
      For j% = 1 To NbPar%
         elem = elem + Wmat(i%, j%) * Vvec(j%)
      Next j%
      Uvec(i%) = elem
   Next i%
   '--------------- Calcul des Upar(i) ----------
   For i% = 1 To NbPar%
      Upar(i%) = UApar(i%) + Uvec(i%)
      UApar(i%) = Upar(i%)
   Next i%
   Exit Sub
Traite_ErreursMoCar:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
         Exit Sub
      Case Else
         MsgBox Error$, 48
         Exit Sub
   End Select
   Exit Sub
End Sub


Private Sub mnuCalculer_Click()
   '--------------------------------------------------------------
   '----- Affectation de leurs valeurs ---------------------------
   '-----  aux variables et paramètres ---------------------------
   '--------------------------------------------------------------
   Call Affecte
   '-------------------------- Variables -------------------------
   NombVar% = 1
   var$(1) = "X"
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   Call Parametre(ForVar2$, ParVar2$)
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Calcul des paramètres de la courbe -----------------
   '--------- pour l'ajuster au nuage de points  -----------------
   '--------------------------------------------------------------
   lblValEr.ForeColor = ROUGE
   lblValEr.Caption = "Calcul en cours..."
   DoEvents
   Call MoinCar
   Call Parametre(ForVar2$, ParVar2$)
   '--------------------------------------------------------------
   '------ Affichage des paramètres de la courbe -----------------
   '---------------- après l'affinement --------------------------
   '--------------------------------------------------------------
   ' **************************************
   FenetreDefMoinCar.GridPar.Col = 1
   For i% = 1 To NbPar%
      FenetreDefMoinCar.GridPar.Row = i%
      FenetreDefMoinCar.GridPar.Text = Format(Upar(i%), "0.000000")
   Next i%
   ' **************************************
   '--------------------------------------------------------------
   '----------- Calcul et affichage de l'erreur ------------------
   '------- E=(somme des [Y(Xmc(i))-Ymc(i)]²)/NbPts --------------
   '--------------------------------------------------------------
   Emc = 0
   For i% = 1 To NbPtsMoinCar%
      X = Xmc(i%)
      vvar$(1) = LTrim$(Format(X, "0.000000"))
      Call variable(ParVar2$, ValVar2$)
      Call Traitement(ValVar2$, SorVar2$)
      Y = CSng(SorVar2$)
      deltaY = Y - Ymc(i%)
      Emc = Emc + deltaY * deltaY
   Next i%
   Emc = Emc / NbPtsMoinCar%
   Emc = Sqr(Emc)
   ' Affichage de E :
   lblValEr.ForeColor = NOIR
   lblValEr.Caption = Format(Emc, "0.000000")
   '--------------------------------------------------------------End Sub
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
   ' et écriture de leurs coordonnées Xmc() et Ymc()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbPtsMoinCar%
   For iloc% = 1 To NbPtsMoinCar%
      Write #1, Xmc(iloc%)
      Write #1, Ymc(iloc%)
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
   Input #1, NbPtsMoinCar%
   ReDim Xmc(1 To NbPtsMoinCar%)
   ReDim Ymc(1 To NbPtsMoinCar%)
   For iloc% = 1 To NbPtsMoinCar%
    Input #1, Xmc(iloc%)
    Input #1, Ymc(iloc%)
   Next iloc%
   Close #1
   '*******************************************************
   '***** actualisation de la grille des points ************
   '********************************************************
   FenetreDefMoinCar.txtNbPts.Text = Format(NbPtsMoinCar%, "0")
   For i% = 1 To NbPtsMoinCar%
      FenetreDefMoinCar.GridPts.Row = i%
      FenetreDefMoinCar.GridPts.Col = 1
      FenetreDefMoinCar.GridPts.Text = Format(Xmc(i%), "0.0000")
      FenetreDefMoinCar.GridPts.Col = 2
      FenetreDefMoinCar.GridPts.Text = Format(Ymc(i%), "0.0000")
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
   FenetreDefMoinCar.Hide
   Fenetre_Prin.Hide
   Fenetre_Val.Hide
   Dessin.Hide
End Sub


Private Sub mnuTracer_Click()
   '-----------------------------------------------------
   '---------      Vérifications avant   ----------------
   '--------- l'ouverture des fenêtres de tracé ---------
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
   If NbPtsMoinCar% < NbPar% Then
      MsgBox "Le nombre de points doit être supérieur ou égal au nombre de paramètres !", 48, "COURBE"
      Exit Sub
   ElseIf AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "COURBE"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "COURBE"
      Exit Sub
   End If
   '-----------------------------------------
   '--------   si quelque chose   -----------
   '--------    a été modifié,    -----------
   '--------   calculer d'abord   -----------
   '-----------------------------------------
   If FenetreDefMoinCar.lblValEr.Caption = "" Then
      mnuCalculer_Click
   End If
   '-----------------------------------------------------
   '--------   fermeture de la fenêtre   ----------------
   '--------      et ouverture           ----------------
   '--------  des fenêtres pour le tracé ----------------
   '-----------------------------------------------------
   Fenetre_Prin.Show
   Fenetre_Val.Show
   Dessin.Show
   '--------------------------------------------------------------
   '            Affichage de la formule
   '             définissant la courbe
   '--------------------------------------------------------------
   Fenetre_Prin.Caption = "Ajustement d'une courbe à un nuage de points"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Détermination par la méthode des moindres carrés des paramètres d'une fonction pour qu'elle soit le mieux ajustée à un nuage de points."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.Caption = "F(X) = " & TextVar2$
   Fenetre_Val.lblNote.ForeColor = CouleurCou&
   TexteNote$ = ""
   For i% = 1 To NbPar%
      TexteNote$ = TexteNote$ & "U(" & i% & ") = " & Format(Upar(i%), "0.0000") & "; "
   Next i%
   Fenetre_Val.lblNote.Caption = TexteNote$
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

Private Sub PointsTracer_Click()
    Trap% = PointsTracer.Value
End Sub

Private Sub txtNbPar_Change()
   NbPar% = CInt(txtNbPar.Text)
   GridPar.Rows = NbPar% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   GridPar.Col = 0
   For i% = 1 To NbPar%
      GridPar.Row = i%
      GridPar.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub

Private Sub txtNbPts_Change()
   NbPtsMoinCar% = CInt(txtNbPts.Text)
   GridPts.Rows = NbPtsMoinCar% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   GridPts.Col = 0
   For i% = 1 To NbPtsMoinCar%
      GridPts.Row = i%
      GridPts.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub




Public Sub Affecte()
   On Error GoTo Traite_ErreursAfDeMoCar
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Xmc, Ymc, Upar, UApar
   ReDim Xmc(1 To NbPtsMoinCar%), Ymc(1 To NbPtsMoinCar%)
   ReDim Upar(1 To NbPar%)
   ReDim UApar(1 To NbPar%)
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbPtsMoinCar%
      FenetreDefMoinCar.GridPts.Row = i%
      FenetreDefMoinCar.GridPts.Col = 1
      If FenetreDefMoinCar.GridPts.Text = "" Then
         FenetreDefMoinCar.GridPts.Text = "0"
         Xmc(i%) = 0
      Else
         Xmc(i%) = CSng(FenetreDefMoinCar.GridPts.Text)
      End If
      FenetreDefMoinCar.GridPts.Col = 2
      If FenetreDefMoinCar.GridPts.Text = "" Then
         FenetreDefMoinCar.GridPts.Text = "0"
         Ymc(i%) = 0
      Else
         Ymc(i%) = CSng(FenetreDefMoinCar.GridPts.Text)
      End If
   Next i%
   ' ********************************************
   ' affectation de leurs valeurs aux paramètres
   ' ********************************************
   FenetreDefMoinCar.GridPar.Col = 1
   For i% = 1 To NbPar%
      FenetreDefMoinCar.GridPar.Row = i%
      If FenetreDefMoinCar.GridPar.Text = "" Then
         FenetreDefMoinCar.GridPar.Text = "0"
         Upar(i%) = 0
         UApar(i%) = 0
      Else
         Upar(i%) = CSng(FenetreDefMoinCar.GridPar.Text)
         UApar(i%) = CSng(FenetreDefMoinCar.GridPar.Text)
      End If
   Next i%
   ' ********************************************
   Exit Sub
Traite_ErreursAfDeMoCar:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48, "FenetreDefMoinCar.Deactivate"
         Exit Sub
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "DEFMOINCAR"
         Exit Sub
   End Select
   Exit Sub
End Sub


