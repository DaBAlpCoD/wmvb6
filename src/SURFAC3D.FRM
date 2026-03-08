VERSION 5.00
Begin VB.Form FenetreSurf3D 
   Caption         =   "Définition de la surface 3D"
   ClientHeight    =   5520
   ClientLeft      =   285
   ClientTop       =   1065
   ClientWidth     =   8730
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
   ScaleHeight     =   5520
   ScaleWidth      =   8730
   Begin VB.TextBox DefZmax 
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
      Left            =   7680
      TabIndex        =   30
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox DefZmin 
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
      Left            =   7680
      TabIndex        =   21
      Top             =   2640
      Width           =   975
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
      TabIndex        =   8
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
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
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
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
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
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
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
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   25
      Top             =   4440
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
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   4440
      Width           =   255
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
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   4440
      Width           =   495
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
      Left            =   480
      TabIndex        =   27
      Top             =   3840
      Width           =   7335
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
      Left            =   4800
      TabIndex        =   5
      Top             =   2880
      Width           =   975
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
      Left            =   1920
      TabIndex        =   4
      Top             =   2880
      Width           =   975
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
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Width           =   975
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
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox chkMemeEchelle 
      Caption         =   "même echelle sur les trois axes"
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
      Top             =   1680
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkPlanXOY 
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
      TabIndex        =   20
      Top             =   840
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Label lblZmax 
      Caption         =   "Zmax"
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
      TabIndex        =   22
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblZmin 
      Caption         =   "Zmin"
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
      TabIndex        =   23
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblAccolZ 
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
      Left            =   6960
      TabIndex        =   36
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblLimitesZ 
      Caption         =   "Limites de l'axe des cotes :"
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
      TabIndex        =   35
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label EtiquetteCoulChoisie 
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
      Left            =   2760
      TabIndex        =   33
      Top             =   4440
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
      Left            =   4560
      TabIndex        =   32
      Top             =   2160
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
      Left            =   240
      TabIndex        =   31
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   " Z="
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
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label EtiquetteDefiniCour 
      Caption         =   "Fonction définissant la courbe :"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblYmax 
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
      Left            =   4320
      TabIndex        =   16
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblXmax 
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
      Left            =   1320
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblYmin 
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
      Left            =   4320
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblLimitesY 
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
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblXmin 
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
      Left            =   1320
      TabIndex        =   11
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblLimitesX 
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
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblAccolY 
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
      Left            =   4080
      TabIndex        =   10
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblAccolX 
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
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
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
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label EtiquetteRemarque 
      Caption         =   $"SURFAC3D.frx":0000
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
      Top             =   1440
      Width           =   4695
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
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblPlanXOY 
      Caption         =   "Plan horizontal :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuTracer 
      Caption         =   "&Tracer..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreSurf3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub chkMemeEchelle_Click()
    Ech% = chkMemeEchelle.Value
    If Ech% = 1 Then
        Ymax = FE * (Xmax - Xmin) + Ymin
        DefYmax.Text = Format(Ymax, "0.000")
    End If
End Sub


Private Sub DefVar2_Change()
    TextVar2$ = DefVar2.Text
End Sub

Private Sub DefXmax_Change()
   On Error Resume Next
   Xmax = CSng(DefXmax.Text)
   If Ech% = 1 Then
       AX = Xmax - Xmin
       AY = FE * AX
       Ymax = AY + Ymin
       DefYmax.Text = Format(Ymax, "0.0000")
   End If
   On Error GoTo 0
End Sub

Private Sub DefXmin_Change()
   On Error Resume Next
   Xmin = CSng(DefXmin.Text)
   If Ech% = 1 Then
      AX = Xmax - Xmin
      AY = FE * AX
      Ymax = AY + Ymin
      DefYmax.Text = Format(Ymax, "0.0000")
   End If
   On Error GoTo 0
End Sub

Private Sub DefYmax_Change()
   On Error Resume Next
   Ymax = CSng(DefYmax.Text)
   If Ech% = 1 Then
      AX = Xmax - Xmin
      AY = FE * AX
      Ymin = Ymax - AY
      DefYmin.Text = Format(Ymin, "0.0000")
   End If
   On Error GoTo 0
End Sub

Private Sub DefYmin_Change()
   On Error Resume Next
   Ymin = CSng(DefYmin.Text)
   If Ech% = 1 Then
      AX = Xmax - Xmin
      AY = FE * AX
      Ymax = AY + Ymin
      DefYmax.Text = Format(Ymax, "0.0000")
   End If
   On Error GoTo 0
End Sub



Private Sub DefZmax_Change()
   On Error Resume Next
   Zmax = CSng(DefZmax.Text)
   On Error GoTo 0
End Sub


Private Sub DefZmin_Change()
   On Error Resume Next
   Zmin = CSng(DefZmin.Text)
   On Error GoTo 0
End Sub


Private Sub Form_Activate()
   chkMemeEchelle.Value = Ech%
   EtiquetteFE.Caption = "     " + Format(FE, "0.000")
   DefXmin.Text = Format(Xmin, "0.0000")
   DefXmax.Text = Format(Xmax, "0.0000")
   DefYmin.Text = Format(Ymin, "0.0000")
   DefYmax.Text = Format(Ymax, "0.0000")
   DefZmin.Text = Format(Zmin, "0.0000")
   DefZmax.Text = Format(Zmax, "0.0000")
   DefVar2.Text = TextVar2$
End Sub

Private Sub Form_Load()
   ' **************************
   Xmin = -1
   Xmax = 1
   Ymin = -1
   Ymax = 1
   Zmin = -1
   Zmax = 1
   ' --------------------------
   ' Col ou selle
   TextVar2$ = "X*X-Y*Y"
   ' --------------------------
   ' Facteur d'échelle
   chkMemeEchelle.Value = 1
   Ech% = 1
   FE = 1
   ' --------------------------
   ' Tracer le plan XOY
   chkPlanXOY.Value = 1
   ' **************************
End Sub



Private Sub mnuQuitter_Click()
   FenetreSurf3D.Hide
End Sub

Private Sub mnuTracer_Click()
   FenetreVisuS3D.lblFormule.Caption = "Z = " & TextVar2$
   FenetreVisuS3D.Show
End Sub

Private Sub pctColorCou_Click(Index As Integer)
    pctColorCourbe.BackColor = pctColorCou(Index).BackColor
    CouleurCou& = pctColorCourbe.BackColor
End Sub

