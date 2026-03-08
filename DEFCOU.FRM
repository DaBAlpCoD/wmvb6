VERSION 5.00
Begin VB.Form FenetreDefCourbe 
   Caption         =   "Définition de la courbe"
   ClientHeight    =   7245
   ClientLeft      =   210
   ClientTop       =   825
   ClientWidth     =   11295
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
   ScaleHeight     =   7245
   ScaleWidth      =   11295
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
      Left            =   6720
      TabIndex        =   85
      Top             =   1800
      Value           =   2  'Grayed
      Width           =   855
   End
   Begin VB.PictureBox pctColorTanChoisie 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   84
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pctColorTan 
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
      Left            =   6960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   82
      Top             =   6600
      Width           =   255
   End
   Begin VB.PictureBox pctColorTan 
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
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   81
      Top             =   6600
      Width           =   255
   End
   Begin VB.PictureBox pctColorTan 
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
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   80
      Top             =   6600
      Width           =   255
   End
   Begin VB.PictureBox pctColorTan 
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
      Left            =   6960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   79
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pctColorTan 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   78
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pctColorTan 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   77
      Top             =   6360
      Width           =   255
   End
   Begin VB.TextBox DefXpar 
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
      Left            =   9480
      TabIndex        =   75
      Text            =   "2"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox TanTracer 
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
      Left            =   8040
      TabIndex        =   72
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox pctColorPtsChoisie 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   68
      Top             =   6600
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   4
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   67
      Top             =   6600
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   66
      Top             =   6600
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   65
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   64
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pctColorPts 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   63
      Top             =   6360
      Width           =   255
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
      Left            =   6720
      TabIndex        =   61
      Top             =   1440
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
      Left            =   6960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   59
      Top             =   5640
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
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   58
      Top             =   5640
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
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   57
      Top             =   5640
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
      Left            =   6960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   5400
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
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   5400
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
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   5400
      Width           =   255
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
      Left            =   5160
      TabIndex        =   52
      Top             =   1440
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
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   51
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox DefTmax 
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
      Left            =   2760
      TabIndex        =   7
      Text            =   "6.3"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox DefTmin 
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
      Left            =   2760
      TabIndex        =   6
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox DefVar2 
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
      TabIndex        =   33
      Text            =   "SIN(X)"
      Top             =   4920
      Width           =   7935
   End
   Begin VB.TextBox DefVar1 
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
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   7935
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
      Left            =   7560
      TabIndex        =   5
      Text            =   "4"
      Top             =   3960
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
      Left            =   2640
      TabIndex        =   4
      Text            =   "4"
      Top             =   3960
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
      Left            =   7560
      TabIndex        =   3
      Text            =   "-4"
      Top             =   3720
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
      Left            =   2640
      TabIndex        =   2
      Text            =   "-4"
      Top             =   3720
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
      Top             =   2880
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
      Height          =   495
      Left            =   2760
      TabIndex        =   28
      Top             =   2040
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
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Value           =   2  'Grayed
      Width           =   2415
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
      Left            =   2760
      TabIndex        =   27
      Top             =   1800
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
      TabIndex        =   23
      Top             =   1800
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
      Left            =   2760
      TabIndex        =   26
      Top             =   1440
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
      TabIndex        =   22
      Top             =   1440
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
      TabIndex        =   45
      Top             =   0
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
         TabIndex        =   47
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
         TabIndex        =   46
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame CadreExpr 
      Caption         =   "Expression"
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
      Left            =   2400
      TabIndex        =   38
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton OptionPar 
         Caption         =   "paramétrée"
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
         TabIndex        =   41
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptionImpl 
         Caption         =   "implicite"
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
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptionExpl 
         Caption         =   "explicite"
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
         TabIndex        =   39
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label EtiquettePar 
         Caption         =   "X=f1(T); Y=f2(T)"
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
         TabIndex        =   44
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label EtiquetteImpl 
         Caption         =   "f(X,Y)=0"
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
         TabIndex        =   43
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label EtiquetteExpl 
         Caption         =   "Y=f(X)"
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
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame CadreCoor 
      Caption         =   "Coordonnées"
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
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton OptionPol 
         Caption         =   "polaires  (TETA,R)"
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
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton OptionCart 
         Caption         =   "cartésiennes (X,Y)"
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
         TabIndex        =   36
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Label EtiquetteCoulTanChoisie 
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
      Left            =   7560
      TabIndex        =   83
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label EtiquetteCoulTan 
      Caption         =   "Couleur de la tangente :"
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
      Left            =   5160
      TabIndex        =   76
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label EtiquetteValPtTan 
      Alignment       =   1  'Right Justify
      Caption         =   "Xo "
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
      Left            =   8160
      TabIndex        =   74
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label EtiquettePtTan 
      Caption         =   "...passant par le point :"
      Height          =   255
      Left            =   8040
      TabIndex        =   73
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label EtiquetteTangente 
      Caption         =   "Tangente :"
      Height          =   255
      Left            =   8040
      TabIndex        =   71
      Top             =   1200
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   69
      Top             =   6360
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
      Left            =   240
      TabIndex        =   62
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label EtiquettePoints 
      Caption         =   "Points :"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   6720
      TabIndex        =   60
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label EtiquetteGrille 
      Caption         =   "Grille :"
      Height          =   255
      Left            =   5160
      TabIndex        =   53
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label EtiquetteCoulCouChoisie 
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
      Left            =   7560
      TabIndex        =   50
      Top             =   5400
      Width           =   855
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
      Left            =   3480
      TabIndex        =   49
      Top             =   3360
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
      Left            =   5400
      TabIndex        =   48
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label EtiquetteTmax 
      Caption         =   "Tmax"
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
      Left            =   2160
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EtiquetteTmin 
      Caption         =   "Tmin"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EtiquetteAccolT 
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label EtiquetteLimitesT 
      Caption         =   "Limites de variation du paramètre :"
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
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "      Y="
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
      Left            =   0
      TabIndex        =   31
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label EtiquetteVar1 
      Caption         =   "      X="
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
      Left            =   0
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label EtiquetteDefiniCour 
      Caption         =   "Fonction(s) définissant la courbe :"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   4320
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
      Left            =   6960
      TabIndex        =   18
      Top             =   3960
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
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
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
      Left            =   6960
      TabIndex        =   16
      Top             =   3720
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
      Left            =   5400
      TabIndex        =   9
      Top             =   3720
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
      Left            =   2040
      TabIndex        =   13
      Top             =   3720
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
      Left            =   360
      TabIndex        =   8
      Top             =   3720
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
      Left            =   6720
      TabIndex        =   12
      Top             =   3600
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
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
      Left            =   4080
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"DEFCOU.frx":0000
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
      TabIndex        =   0
      Top             =   2640
      Width           =   5535
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
      TabIndex        =   34
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label EtiquetteAxes 
      Caption         =   "Axes :"
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label EtiquetteCadre 
      Caption         =   "Cadre :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1200
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
Attribute VB_Name = "FenetreDefCourbe"
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
        CadreAfficher.ForeColor = &H80000011
    ElseIf Grac% = 1 Then
        CadreAfficher.Value = 0
        CadreAfficher.ForeColor = &H80000012
    End If
End Sub

Private Sub CadreTracer_Click()
    Trac% = CadreTracer.Value
End Sub



Private Sub DefTmax_Change()
   On Error Resume Next
   Tmax = CSng(DefTmax.Text)
   On Error GoTo 0
End Sub

Private Sub DefTmin_Change()
   On Error Resume Next
   Tmin = CSng(DefTmin.Text)
   On Error GoTo 0
End Sub

Private Sub DefVar1_Change()
    TextVar1$ = DefVar1.Text
End Sub

Private Sub DefVar2_Change()
    TextVar2$ = DefVar2.Text
End Sub




Private Sub DefXpar_Change()
    Xpar = CSng(DefXpar.Text)
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
   On Error Resume Next
   Select Case Gap%
      Case 1
         Xmin = CSng(DefXmin.Text)
         If Ech% = 1 Then
            AX = Xmax - Xmin
            AY = FE * AX
            Ymax = AY + Ymin
            DefYmax.Text = Format(Ymax, "0.0000")
         End If
      Case 2
         DefXmin.Text = Format(Xmin, "0.0000")
   End Select
   On Error GoTo 0
End Sub

Private Sub DefYmax_Change()
   On Error Resume Next
   Select Case Gap%
      Case 1
         Ymax = CSng(DefYmax.Text)
         If Ech% = 1 Then
            AX = Xmax - Xmin
            AY = FE * AX
            Ymin = Ymax - AY
            DefYmin.Text = Format(Ymin, "0.0000")
         End If
      Case 2
         DefYmax.Text = Format(Ymax, "0.0000")
   End Select
   On Error GoTo 0
End Sub

Private Sub DefYmin_Change()
   On Error Resume Next
   Select Case Gap%
      Case 1
         Ymin = CSng(DefYmin.Text)
         If Ech% = 1 Then
            AX = Xmax - Xmin
            AY = FE * AX
            Ymax = AY + Ymin
            DefYmax.Text = Format(Ymax, "0.0000")
         End If
      Case 2
         DefYmin.Text = Format(Ymin, "0.0000")
   End Select
   On Error GoTo 0
End Sub




Private Sub Form_Activate()
   Select Case Coor%
   Case 1
      OptionCart.Value = -1
      EtiquetteExpl.Caption = "Y=f(X)"
      EtiquetteImpl.Caption = "f(X,Y)=0"
      EtiquettePar.Caption = "X=f1(T); Y=f2(T)"
      EtiquetteVar1.Caption = "   X="
      EtiquetteVar2.Caption = "   Y="
   Case 2
      OptionPol.Value = -1
      EtiquetteExpl.Caption = "R=f(TETA)"
      EtiquetteImpl.Caption = "f(TETA,R)=0"
      EtiquettePar.Caption = "TETA=f1(T); R=f2(T)"
      EtiquetteVar1.Caption = "TETA="
      EtiquetteVar2.Caption = "   R="
   End Select
   Select Case Expr%
   Case 1
      OptionExpl.Value = -1
      EtiquetteLimitesT.Visible = 0
      EtiquetteAccolT.Visible = 0
      EtiquetteTmin.Visible = 0
      EtiquetteTmax.Visible = 0
      DefTmin.Visible = 0
      DefTmax.Visible = 0
      EtiquetteVar1.Visible = 0
      DefVar1.Visible = 0
   Case 2
      OptionImpl.Value = -1
      EtiquetteLimitesT.Visible = 0
      EtiquetteAccolT.Visible = 0
      EtiquetteTmin.Visible = 0
      EtiquetteTmax.Visible = 0
      DefTmin.Visible = 0
      DefTmax.Visible = 0
      EtiquetteVar1.Visible = 0
      DefVar1.Visible = 0
   Case 3
      OptionPar.Value = -1
      EtiquetteLimitesT.Visible = -1
      EtiquetteAccolT.Visible = -1
      EtiquetteTmin.Visible = -1
      EtiquetteTmax.Visible = -1
      DefTmin.Visible = -1
      DefTmax.Visible = -1
      EtiquetteVar1.Visible = -1
      DefVar1.Visible = -1
   End Select
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
   DefXpar.Text = Format(Xpar, "0.0000")
   DefXmin.Text = Format(Xmin, "0.0000")
   DefXmax.Text = Format(Xmax, "0.0000")
   DefYmin.Text = Format(Ymin, "0.0000")
   DefYmax.Text = Format(Ymax, "0.0000")
   DefTmin.Text = Format(Tmin, "0.0000")
   DefTmax.Text = Format(Tmax, "0.0000")
   DefVar1.Text = TextVar1$
   DefVar2.Text = TextVar2$
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
   Rep% = 0
   Vax% = 1
   Tran% = 0
   Ech% = 1
   Xpar = 1
   Xmin = -4
   Xmax = 4
   Ymin = -4
   Ymax = 4
   TextVar2$ = "SIN(X)"
   CouleurCou& = ROUGE
   CouleurPts& = BLEU
   CouleurTan& = VERT
   '-----------------------------------------------------------------
   '-----------        Valeurs par défaut             ---------------
   '-----------        d'autres variables             ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   AX = Xmax - Xmin
   AY = Ymax - Ymin
   MX = LX% / AX
   MY = -LY% / AY
End Sub

Private Sub GrilleTracer_Click()
    Tric% = GrilleTracer.Value
End Sub

Private Sub Label1_Click()

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
   FenetreDefCourbe.Hide
   Fenetre_Prin.Hide
   Fenetre_Val.Hide
   Dessin.Hide
End Sub

Private Sub mnuTracer_Click()
   '-----------------------------------------------------
   '---------   Vérifications avant   -------------------
   '---------------   tracé   ---------------------------
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
   If Expr% = 3 Then
      Tmax = CSng(DefTmax.Text)
      If Err.Number <> 0 Then
         MsgBox "Tmax est incorrect", 48, "COURBE"
         Exit Sub
      End If
      Tmin = CSng(DefTmin.Text)
      If Err.Number <> 0 Then
         MsgBox "Tmin est incorrect", 48, "COURBE"
         Exit Sub
      End If
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
      MsgBox "Xmax doit être supérieur à Xmin", 48, "COURBE"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "COURBE"
      Exit Sub
   ElseIf Tmax < Tmin Then
      MsgBox "Tmax doit être supérieur à Tmin", 48, "COURBE"
      Exit Sub
   End If
   On Error GoTo 0
   '-----------------------------------------------------
   '--------         Ouverture           ----------------
   '--------  des fenêtres pour le tracé ----------------
   '-----------------------------------------------------
   Fenetre_Prin.Show
   Fenetre_Val.Show
   Dessin.Show
   '--------------------------------------------------------------
   '--------- Affichage de la (des) formule(s) -------------------
   '---------     définissant la courbe        -------------------
   '--------------------------------------------------------------
   Fenetre_Prin.Caption = "Courbe"
   Fenetre_Prin.LabelExplications.ForeColor = MAGENTA
   Fenetre_Prin.LabelExplications.Caption = "Courbe en coordonnées cartésiennes ou polaires; définie de manière explicite, implicite, ou paramétrée."
   Fenetre_Val.Cls
   Fenetre_Val.ForeColor = CouleurCou&
   Fenetre_Val.lblFormule.ForeColor = CouleurCou&
   Select Case Expr%
   Case 1
      Select Case Coor%
      Case 1
         Fenetre_Val.lblFormule.Caption = " Y=" & TextVar2$
      Case 2
         Fenetre_Val.lblFormule.Caption = " R=" & TextVar2$
      End Select
   Case 2
      Fenetre_Val.lblFormule.Caption = TextVar2$ & "=0"
   Case 3
      Select Case Coor%
      Case 1
         Fenetre_Val.lblFormule.Caption = "X=" & TextVar1$ & Chr$(13) & "Y=" & TextVar2$
      Case 2
         Fenetre_Val.lblFormule.Caption = "TETA=" & TextVar1$ & Chr$(13) & "R=" & TextVar2$
      End Select
      Fenetre_Val.lblNote.ForeColor = CouleurCou&
      Fenetre_Val.lblNote.Caption = "Tmin=" & Tmin & Chr$(13) & " Tmax=" & Tmax
   End Select
   '------------------------------------------------
   '--------- Lancement du tracé -------------------
   '--------------------------------------------------------------
   Rem Fenetre_Prin.mnuTraceStop_Click
   Rem ne trace pas les axes !
End Sub

Private Sub OptionCart_Click()
    Coor% = 1
    EtiquetteExpl.Caption = "Y=f(X)"
    EtiquetteImpl.Caption = "f(X,Y)=0"
    EtiquettePar.Caption = "X=f1(T); Y=f2(T)"
    EtiquetteVar1.Caption = "   X="
    EtiquetteVar2.Caption = "   Y="
    EtiquetteValPtTan.Caption = "Xo"
    If Expr% = 1 Then
        DefVar2.Text = "SIN(X)"
    ElseIf Expr% = 2 Then
        EtiquetteVar2.Caption = "   0="
        DefVar2.Text = "X*X/9+Y*Y/4-1"
    ElseIf Expr% = 3 Then
        DefVar1.Text = "4 * COS(3 * T)"
        DefVar2.Text = "4 * SIN(4 * T)"
        DefTmin.Text = "0"
        DefTmax.Text = "6.3"
    End If
End Sub

Private Sub OptionEffacer_Click()
    Gap% = 1
End Sub

Private Sub OptionExpl_Click()
    Expr% = 1
    EtiquetteLimitesT.Visible = 0
    EtiquetteAccolT.Visible = 0
    EtiquetteTmin.Visible = 0
    EtiquetteTmax.Visible = 0
    DefTmin.Visible = 0
    DefTmax.Visible = 0
    EtiquetteVar1.Visible = 0
    DefVar1.Visible = 0
    If Coor% = 1 Then
        EtiquetteVar2.Caption = "   Y="
        DefVar2.Text = "SIN(X)"
    Else
        EtiquetteVar2.Caption = "   R="
        DefVar2.Text = "3*COS(TETA)+1"
    End If
End Sub

Private Sub OptionGarder_Click()
    Gap% = 2
End Sub

Private Sub OptionImpl_Click()
    Expr% = 2
    EtiquetteLimitesT.Visible = 0
    EtiquetteAccolT.Visible = 0
    EtiquetteTmin.Visible = 0
    EtiquetteTmax.Visible = 0
    DefTmin.Visible = 0
    DefTmax.Visible = 0
    EtiquetteVar1.Visible = 0
    DefVar1.Visible = 0
    EtiquetteVar2.Caption = "   0="
    If Coor% = 1 Then
        DefVar2.Text = "X*X/16+Y*Y/9-1"
    Else
        DefVar2.Text = "R-4*ABS(COS(4*TETA)"
    End If
End Sub

Private Sub OptionPar_Click()
    Expr% = 3
    EtiquetteValPtTan.Caption = "To"
    EtiquetteLimitesT.Visible = -1
    EtiquetteAccolT.Visible = -1
    EtiquetteTmin.Visible = -1
    EtiquetteTmax.Visible = -1
    DefTmin.Visible = -1
    DefTmax.Visible = -1
    EtiquetteVar1.Visible = -1
    DefVar1.Visible = -1
    If Coor% = 1 Then
        Tmin = 0
        Tmax = 6.3
        EtiquetteVar2.Caption = "   Y="
        DefVar1.Text = "4 * COS(3 * T)"
        DefVar2.Text = "4 * SIN(4 * T)"
        DefTmin.Text = Format(Tmin, "0.0000")
        DefTmax.Text = Format(Tmax, "0.0000")
    Else
        Tmin = 1
        Tmax = 30
        EtiquetteVar2.Caption = "   R="
        DefVar1.Text = "T"
        DefVar2.Text = "LOG(T)"
        DefTmin.Text = Format(Tmin, "0.0000")
        DefTmax.Text = Format(Tmax, "0.0000")
    End If
End Sub

Private Sub OptionPol_Click()
    Coor% = 2
    EtiquetteExpl.Caption = "R=f(TETA)"
    EtiquetteImpl.Caption = "f(TETA,R)=0"
    EtiquettePar.Caption = "TETA=f1(T); R=f2(T)"
    EtiquetteVar1.Caption = "TETA="
    EtiquetteVar2.Caption = "   R="
    EtiquetteValPtTan.Caption = "TETAo"
    If Expr% = 1 Then
        DefVar2.Text = "2*COS(TETA)+1"
    ElseIf Expr% = 2 Then
        EtiquetteVar2.Caption = "   0="
        DefVar2.Text = "R-4*ABS(COS(4*TETA))"
    ElseIf Expr% = 3 Then
        DefVar1.Text = "T"
        DefVar2.Text = "LOG(T)"
        DefTmin.Text = "1"
        DefTmax.Text = "30"
    End If
End Sub

Private Sub pctColorCou_Click(Index As Integer)
   pctColorCourbe.BackColor = pctColorCou(Index).BackColor
   CouleurCou& = pctColorCourbe.BackColor
End Sub


Private Sub pctColorPts_Click(Index As Integer)
   pctColorPtsChoisie.BackColor = pctColorPts(Index).BackColor
   CouleurPts& = pctColorPtsChoisie.BackColor
End Sub

Private Sub pctColorTan_Click(Index As Integer)
   pctColorTanChoisie.BackColor = pctColorTan(Index).BackColor
   CouleurTan& = pctColorTanChoisie.BackColor
End Sub

Private Sub PointsRelier_Click()
    Rep% = PointsRelier.Value
End Sub

Private Sub PointsTracer_Click()
    Trap% = PointsTracer.Value
End Sub

Private Sub TanTracer_Click()
    Tran% = TanTracer.Value
End Sub
