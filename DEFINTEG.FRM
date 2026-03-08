VERSION 5.00
Begin VB.Form FenetreDefInteg 
   Caption         =   "Définition de la courbe [intégrale d'une fonction]"
   ClientHeight    =   7200
   ClientLeft      =   345
   ClientTop       =   735
   ClientWidth     =   9885
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
   ScaleHeight     =   7200
   ScaleWidth      =   9885
   Begin VB.PictureBox pctColorPtsChoisie 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   68
      Top             =   6720
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   4
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   67
      Top             =   6720
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   66
      Top             =   6720
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   65
      Top             =   6480
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   64
      Top             =   6480
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   63
      Top             =   6480
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
      Left            =   5760
      TabIndex        =   61
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
      Left            =   5760
      TabIndex        =   60
      Top             =   360
      Value           =   2  'Grayed
      Width           =   855
   End
   Begin VB.PictureBox pctColorZ 
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
      Left            =   6600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox pctColorZ 
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
      Left            =   6360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox pctColorZ 
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
      Left            =   6120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox pctColorZ 
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
      Left            =   6600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Top             =   5640
      Width           =   255
   End
   Begin VB.PictureBox pctColorZ 
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
      Left            =   6360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   5640
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
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   5880
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
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Top             =   5880
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
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   5880
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
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   5640
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
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   48
      Top             =   5640
      Width           =   255
   End
   Begin VB.PictureBox pctColorZone 
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
      Height          =   495
      Left            =   7920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   5640
      Width           =   495
   End
   Begin VB.PictureBox pctColorZ 
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
      Left            =   6120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   45
      Top             =   5640
      Width           =   255
   End
   Begin VB.TextBox txtXb 
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
      Left            =   720
      TabIndex        =   41
      Text            =   "1"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtXa 
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
      Left            =   720
      TabIndex        =   39
      Text            =   "0"
      Top             =   4920
      Width           =   1695
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
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   5640
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
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   36
      Top             =   5640
      Width           =   255
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
      Left            =   960
      TabIndex        =   27
      Text            =   "X"
      Top             =   4200
      Width           =   7455
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
      Left            =   6840
      TabIndex        =   5
      Text            =   "4"
      Top             =   3240
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
      Top             =   3240
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
      Left            =   6840
      TabIndex        =   3
      Text            =   "-4"
      Top             =   3000
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
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox MemeEchelle 
      Caption         =   "même echelle sur les deux axes"
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
      Top             =   2040
      Value           =   1  'Checked
      Width           =   3015
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
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1080
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
      TabIndex        =   31
      Top             =   240
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
      Left            =   2400
      TabIndex        =   69
      Top             =   6480
      Width           =   855
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
      Left            =   120
      TabIndex        =   62
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label EtiquettePoints 
      Caption         =   "Points :"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   5760
      TabIndex        =   59
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMethode 
      Caption         =   "Calcul de l'intégrale :  I =  somme de Xa à Xb de f(t) dt  par la méthode de SIMPSON."
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
      Left            =   2760
      TabIndex        =   58
      Top             =   4680
      Width           =   4695
   End
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"DEFINTEG.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   57
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label lblCoulInt 
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
      Left            =   7080
      TabIndex        =   47
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label lblZoneInt 
      Caption         =   "Couleur de la zone intégrée :"
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
      TabIndex        =   44
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label lblInteg 
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
      Left            =   3480
      TabIndex        =   43
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Label lblI 
      Caption         =   "I ="
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
      TabIndex        =   42
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblXb 
      Caption         =   "Xb "
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
      TabIndex        =   40
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblXa 
      Caption         =   "Xa "
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
      TabIndex        =   38
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lblBornes 
      Caption         =   "Bornes d'intégration :"
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
      TabIndex        =   37
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label EtiquetteGrille 
      Caption         =   "Grille :"
      Height          =   255
      Left            =   4680
      TabIndex        =   29
      Top             =   120
      Width           =   735
   End
   Begin VB.Label EtiquetteCoulCou 
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
      Left            =   2400
      TabIndex        =   30
      Top             =   5640
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
      Top             =   2520
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
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "f(X) ="
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
      Left            =   240
      TabIndex        =   26
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblFonction 
      Caption         =   "Fonction à intégrer :"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   2055
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
      Left            =   6240
      TabIndex        =   15
      Top             =   3240
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
      Top             =   3240
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
      Left            =   6240
      TabIndex        =   14
      Top             =   3000
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
      Left            =   4560
      TabIndex        =   7
      Top             =   3000
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
      Top             =   3000
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
      Top             =   3000
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
      Left            =   6000
      TabIndex        =   10
      Top             =   2880
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
      Top             =   2880
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
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label EtiquetteAccolRem 
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   42
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   3120
      TabIndex        =   28
      Top             =   1800
      Width           =   375
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
         Caption         =   "&Ouvrir"
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
Attribute VB_Name = "FenetreDefInteg"
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


Private Sub CommandeOK_Click()
   AX = Xmax - Xmin
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
   End If
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "COURBE"
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "COURBE"
   ElseIf Tmax < Tmin Then
      MsgBox "Tmax doit être supérieur à Tmin", 48, "COURBE"
   Else
      FenetreDefInteg.Hide
   End If
End Sub

Private Sub DefVarf_Change()
   TextVar2$ = DefVarf.Text
   FenetreIntegrale.DefVarf.Text = TextVar2$
   FenetreDefInteg.lblInteg.Caption = ""
End Sub

Private Sub DefXmax_Change()
   On Error Resume Next
   Select Case Gap%
      Case 1
         Xmax = CSng(DefXmax.Text)
         If Ech% = 1 Then
            AX = Xmax - Xmin
            AY = FE * AX
            Ymax = AY + Ymin
            DefYmax.Text = Format(Ymax, "0.0000")
         End If
      Case 2
         DefXmax.Text = Format(Xmax, "0.0000")
   End Select
   On Error GoTo 0
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
   '*****************************************
   '***** actualisations des paramètres *****
   '*****        pour le tracé          *****
   '*****************************************
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
   '---------------------------------------------------------------
   CouleurCou& = FenetreDefInteg.pctColorCourbe.BackColor
   CouleurHach& = FenetreDefInteg.pctColorZone.BackColor
   '****************************************
   '*****  actualisations spécifiques  *****
   '*****       à cette fenêtre        *****
   '****************************************
   FenetreDefInteg.txtXa.Text = Format(Xa, "0.000000")
   FenetreDefInteg.txtXb.Text = Format(Xb, "0.000000")
   FenetreDefInteg.DefVarf.Text = TextVar2$
   FenetreDefInteg.lblInteg.Caption = Format(Integrale, "0.0000")
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
   'TextVar2$ = "X"
   ' La primitive de f(X) = X est X²/2 et
   ' L'intégrale I de f(X) entre 0 et 1
   ' est : I = 0.5
   CouleurCou& = ROUGE
   CouleurPts& = BLEU
   CouleurHach& = BLEU
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
   If FenetreDefInteg.lblInteg.Caption = "" Then
      DoEvents
      '--------------------------------------------------------------
      '--------- Première manipulation des chaines DefVar : ---------
      '--------- Suppression des blancs dans les formules,  ---------
      '--------- vérification du nombre de parenthèses et   ---------
      '--------- remplacement des constantes par leur valeur --------
      '--------------------------------------------------------------
      NombVar% = 1
      var$(1) = "X"
      TextVar2$ = UCase$(TextVar2$)
      Call OteBlancs(TextVar2$, CorVar2$)
      Call Constante(CorVar2$, ForVar2$)
      If Erreur = True Then
         MsgBox "Erreur dans la formule !", 48, "DefInteg"
         Exit Sub
      End If
      '--------------------------------------------------------------
      Call Simpson
      '--------------------------------------------------------------
      FenetreDefInteg.lblInteg.ForeColor = NOIR
      FenetreDefInteg.lblInteg.Caption = Format(Integrale, "0.0000")
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
   FenetreDefInteg.Hide
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
   Xa = CSng(txtXa.Text)
   If Err.Number <> 0 Then
      MsgBox "Xa est incorrect", 48, "COURBE"
      Exit Sub
   End If
   Xb = CSng(txtXb.Text)
   If Err.Number <> 0 Then
      MsgBox "Xb est incorrect", 48, "COURBE"
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
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "FenetreDefInteg.mnuTracer"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "FenetreDefInteg.mnuTracer"
      Exit Sub
   End If
   '-----------------------------------------------------
   '--------   si la formule de la fonction   -----------
   '--------      f(X) a été modifiée,        -------------
   '--------       calculer d'abord           -----------
   '-----------------------------------------------------
   If FenetreDefInteg.lblInteg.Caption = "" Then
      mnuCalculer_Click
   End If
   '-----------------------------------------------------
   '--------   fermeture de la fenêtre   ----------------
   '--------      et ouverture           ----------------
   '--------  des fenêtres pour le tracé ----------------
   '-----------------------------------------------------
   FenetreDefInteg.Hide
   Fenetre_Prin.Show
   Fenetre_Val.Show
   Dessin.Show
   '--------------------------------------------------------------
   '--------- Affichage de la formule ----------------------------
   '---------     définissant la courbe        -------------------
   '--------------------------------------------------------------
   Fenetre_Prin.Caption = "Intégrale d'une fonction"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Intégrale = somme de Xa à Xb de f(t) dt" & Chr$(13) & "Résolution par la méthode de SIMPSON."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.Caption = "f(X) = " & TextVar2$
   Fenetre_Val.lblNote.ForeColor = CouleurHach&
   Fenetre_Val.lblNote.Caption = " Intégrale entre " & Format(Xa, "0.00") & " et " & Format(Xb, "0.00") & " de f(X) :" & Chr$(13) & " I = " & Format(Integrale, "0.00")
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

Private Sub pctColorZ_Click(Index As Integer)
   pctColorZone.BackColor = pctColorZ(Index).BackColor
   CouleurHach& = pctColorZone.BackColor
End Sub

Private Sub PointsRelier_Click()
    Rep% = PointsRelier.Value
End Sub

Private Sub PointsTracer_Click()
    Trap% = PointsTracer.Value
End Sub

Private Sub txtXa_Change()
   On Error Resume Next
   Xa = CSng(txtXa.Text)
   lblInteg.Caption = ""
   On Error GoTo 0
End Sub

Private Sub txtXb_Change()
   On Error Resume Next
   Xb = CSng(txtXb.Text)
   lblInteg.Caption = ""
   On Error GoTo 0
End Sub

