VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreTriangle 
   Caption         =   "Triangles"
   ClientHeight    =   5835
   ClientLeft      =   420
   ClientTop       =   945
   ClientWidth     =   8775
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
   ScaleHeight     =   5835
   ScaleWidth      =   8775
   Begin MSFlexGridLib.MSFlexGrid GridCentres 
      Height          =   1935
      Left            =   4560
      TabIndex        =   0
      Top             =   3840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   6
      Cols            =   4
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
   Begin MSFlexGridLib.MSFlexGrid GridMHBM 
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   5
      Cols            =   4
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
   Begin MSFlexGridLib.MSFlexGrid GridAngCot 
      Height          =   1095
      Left            =   3120
      TabIndex        =   13
      Top             =   2520
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   4
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
   Begin MSFlexGridLib.MSFlexGrid GridSom 
      Height          =   1215
      Left            =   3120
      TabIndex        =   14
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   4
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
   Begin VB.PictureBox pctTriangle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
      Begin VB.Label lblcotea 
         Caption         =   "a"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblcoteb 
         Caption         =   "b"
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
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblcotec 
         Caption         =   "c"
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
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblC 
         Caption         =   "C"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblB 
         Caption         =   "B"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblA 
         Caption         =   "A"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line3 
         X1              =   480
         X2              =   2160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   2160
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   480
         Y1              =   120
         Y2              =   1200
      End
   End
   Begin VB.Label lblCaracTriangle 
      Caption         =   "quelconque avec trois angles aigus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblTypeTriangle 
      Caption         =   "Triangle exemple :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblValDemiP 
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
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblValSurf 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblSurface 
      Caption         =   "Surface :"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblDemiPer 
      Caption         =   "Demi-périmètre :"
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
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuTriangle 
      Caption         =   "&Triangle"
      Begin VB.Menu mnuTriangleEx 
         Caption         =   "Triangle &exemple"
         Begin VB.Menu mnuTroisAnglesAigus 
            Caption         =   "Quelc&onque avec trois angles aigus"
         End
         Begin VB.Menu mnuUnAngleObtus 
            Caption         =   "Quelconque avec &un angle obtus"
         End
         Begin VB.Menu mnuRectangle 
            Caption         =   "&Rectangle"
         End
         Begin VB.Menu mnuIsocele 
            Caption         =   "&Isocèle"
         End
         Begin VB.Menu mnuRectIso 
            Caption         =   "Rectang&le isocèle"
         End
         Begin VB.Menu mnuEquilateral 
            Caption         =   "Equil&atéral"
         End
      End
      Begin VB.Menu mnuTriangleDef 
         Caption         =   "Triangle à &définir"
         Begin VB.Menu mnuDonneesCoor 
            Caption         =   "Données : coordonnées des &sommets"
         End
         Begin VB.Menu mnuDonneesCotAng 
            Caption         =   "Données : côtés et a&ngles"
         End
      End
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer"
   End
   Begin VB.Menu mnuVisualiser 
      Caption         =   "&Visualiser..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------
' Drapeau : égal à FALSE si les calculs
' n'ont pas été fait depuis le dernier changement,
' à TRUE sinon.
Dim TrianFlag As Boolean
' ------------------------------------------------
' Erreur dans les calculs : FALSE si pas d'erreur,
' TRUE si erreur.
Dim ErrTrian As Boolean
' ------------------------------------------------

Private Sub AngCotVersSom()
' ----------------------------------------------------
' Calcule les sommets à partir des angles et des côtés
' ----------------------------------------------------
' coordonnées des sommets
'Global Xtrian(1 To 3), Ytrian(1 To 3)
' différences de coordonnées des sommets
'Global DXtrian(1 To 3), DYtrian(1 To 3)
' angles (en radians) et sommets
'Global AngleTrian(1 To 3), Cote(1 To 3)
'Global Const CONVDEGRAD = 180/PI
' ----------------------------------------------------
Xtrian(1) = Cote(3) * Cos(AngleTrian(2))
Ytrian(1) = Cote(3) * Sin(AngleTrian(2))
Xtrian(2) = 0
Ytrian(2) = 0
Xtrian(3) = Cote(1)
Ytrian(3) = 0
For i% = 1 To 3
   j% = i% + 1
   If j% = 4 Then j% = 1
   k% = j% + 1
   If k% = 4 Then k% = 1
   DXtrian(i%) = Xtrian(k%) - Xtrian(j%)
   DYtrian(i%) = Ytrian(k%) - Ytrian(j%)
Next i%
Cote(2) = Sqr(DXtrian(2) ^ 2 + DYtrian(2) ^ 2)
' demi-périmètre :
DemiPerTrian = (Cote(1) + Cote(2) + Cote(3)) / 2
If DemiPerTrian <= 0 Then
   ErrTrian = True
   Beep
   MsgBox "Erreur dans les données !", 48, "TRIANGLE"
   Exit Sub
End If
' aire
AireTrian = 0
' rayon du cercle circonscrit (médiatrices)
RayonCirc = 1
For i% = 1 To 3
   DemiPerMoinsCote(i%) = DemiPerTrian - Cote(i%)
   AireTrian = AireTrian - Xtrian(i%) * DYtrian(i%)
   RayonCirc = RayonCirc * Cote(i%)
Next i%
AireTrian = AireTrian / 2
If AireTrian <= 0 Then
   AireTrian = -AireTrian
   ' A, B et C ne se succèdent pas
   ' dans le sens trigonométrique
   MsgBox "A, B et C ne se succèdent pas dans le sens trigonométrique", 64, "TRIANGLE"
End If
RayonInscrit = AireTrian / DemiPerTrian
RayonCirc = RayonCirc / AireTrian / 4
RayonEuler = RayonCirc / 2
For i% = 1 To 3
   TangDemiAng = RayonInscrit / DemiPerMoinsCote(i%)
   If i% <> 2 Then
      AngleTrian(i%) = 2 * Atn(TangDemiAng)
   End If
   TanP = 1 - TangDemiAng * TangDemiAng
   TanM = 1 + TangDemiAng * TangDemiAng
   SinAng(i%) = Sin(AngleTrian(i%))
   CosAng(i%) = Cos(AngleTrian(i%))
Next i%
' ***************************************************
'           Affichage sommets
' ***************************************************
For i% = 1 To 3
   GridSom.Row = i%
   GridSom.Col = 1
   GridSom.Text = Format(Xtrian(i%), "0.000")
   GridSom.Col = 2
   GridSom.Text = Format(Ytrian(i%), "0.000")
Next i%
' ***************************************************
'           Affichage angles et côtés
' ***************************************************
For i% = 1 To 3
   GridAngCot.Row = i%
   GridAngCot.Col = 1
   GridAngCot.Text = Format(AngleTrian(i%) * CONVDEGRAD, "000.00")
   GridAngCot.Col = 2
   GridAngCot.Text = Format(Cote(i%), "0.000")
Next i%
' ***************************************************
End Sub

Private Sub CalculsDivers()
Obtus% = 0
For i% = 1 To 3
   If CosAng(i%) < 0 Then Obtus% = i%
   j% = i% + 1
   If j% = 4 Then j% = 1
   k% = j% + 1
   If k% = 4 Then k% = 1
   Xloc = Xtrian(i%)
   Yloc = Ytrian(i%)
   ' ****** Médiane ******
   Xloc1 = (Xtrian(j%) + Xtrian(k%)) / 2
   Yloc1 = (Ytrian(j%) + Ytrian(k%)) / 2
   Xmediane(i%) = Xloc1
   Ymediane(i%) = Yloc1
   Lmediane(i%) = Sqr((Xloc - Xloc1) ^ 2 + (Yloc - Yloc1) ^ 2)
   ' ****** Hauteur ******
   Dloc1 = DXtrian(i%) ^ 2
   Dloc2 = DYtrian(i%) ^ 2
   Dloc3 = Dloc1 + Dloc2
   Xhauteur(i%) = (Xloc * Dloc1 + Xtrian(j%) * Dloc2 - DXtrian(i%) * DYtrian(i%) * DYtrian(k%)) / Dloc3
   Yhauteur(i%) = (Yloc * Dloc2 + Ytrian(j%) * Dloc1 - DXtrian(i%) * DYtrian(i%) * DXtrian(k%)) / Dloc3
   Lhauteur(i%) = 2 * AireTrian / Cote(i%)
   ' ****** Bissectrice ******
   lloc = Cote(k%) / (Cote(j%) + Cote(k%))
   Xloc2 = Xtrian(j%) + lloc * DXtrian(i%)
   Yloc2 = Ytrian(j%) + lloc * DYtrian(i%)
   Xbissec(i%) = Xloc2
   Ybissec(i%) = Yloc2
   Lbissec(i%) = Sqr((Xloc - Xloc2) ^ 2 + (Yloc - Yloc2) ^ 2)
   ' ****** Médiatrice ******
   If Cote(j%) >= Cote(k%) Then
      Rloc1 = DXtrian(j%)
      Rloc2 = DYtrian(j%)
   Else
      Rloc1 = DXtrian(k%)
      Rloc2 = DYtrian(k%)
   End If
   'Lloc = DXtrian(i%) * Rloc1 + DYtrian(i%) * Rloc2
   'Xloc3 = (DXtrian(i%) * Xloc1 * Rloc1 + DYtrian(i%) * (DYtrian(k%) - DYtrian(j%)) / 2 * Rloc1 + Rloc2 * Xtrian(i%)) / Lloc
   'Yloc3 = (DYtrian(i%) * Yloc1 * Rloc2 + DXtrian(i%) * (DXtrian(k%) - DXtrian(j%)) / 2 * Rloc2 + Rloc1 * Ytrian(i%)) / Lloc
   'Xmediat(i%) = Xloc3
   'Ymediat(i%) = Yloc3
   'Lmediat(i%) = Sqr((Xloc3 - Xloc1) ^ 2 + (Yloc3 - Yloc1) ^ 2)
   lloc = DXtrian(i%) * Rloc1 + DYtrian(i%) * Rloc2
   Xloc3 = Xtrian(i%) + Rloc1 * (DXtrian(i%) * (DXtrian(k%) - DXtrian(j%)) + DYtrian(i%) * (DYtrian(k%) - DYtrian(j%))) / lloc / 2
   Yloc3 = Ytrian(i%) + Rloc2 * (DXtrian(i%) * (DXtrian(k%) - DXtrian(j%)) + DYtrian(i%) * (DYtrian(k%) - DYtrian(j%))) / lloc / 2
   Xmediat(i%) = Xloc3
   Ymediat(i%) = Yloc3
   Lmediat(i%) = Sqr((Xloc3 - Xloc1) ^ 2 + (Yloc3 - Yloc1) ^ 2)
Next i%
' ***************************************************************
' ************* Centre de gravité (ou Barycentre) ***************
' ************* point d'intersection des médianes ***************
' ***************************************************************
Xbar = (Xtrian(1) + Xtrian(2) + Xtrian(3)) / 3
Ybar = (Ytrian(1) + Ytrian(2) + Ytrian(3)) / 3
' ***************************************************************
' *************           Orthocentre             ***************
' ************* point d'intersection des hauteurs ***************
' ***************************************************************
lloc = DXtrian(1) * DYtrian(2) - DXtrian(2) * DYtrian(1)
Xort = (Xtrian(1) * DXtrian(1) * DYtrian(2) - DYtrian(1) * DYtrian(2) * DYtrian(3) - Xtrian(2) * DXtrian(2) * DYtrian(1)) / lloc
Yort = -(Ytrian(1) * DYtrian(1) * DXtrian(2) - DXtrian(1) * DXtrian(2) * DXtrian(3) - Ytrian(2) * DYtrian(2) * DXtrian(1)) / lloc
' *******************************************************************
' *************       Centre du cercle inscrit        ***************
' ************* point d'intersection des bissectrices ***************
' *******************************************************************
Lloc0 = DXtrian(3) * DYtrian(2) - DXtrian(2) * DYtrian(3)
Xins = Xtrian(1) + (Cote(3) * DYtrian(2) + Cote(2) * DYtrian(3)) * DemiPerMoinsCote(1) / Lloc0
Yins = Ytrian(1) - (Cote(3) * DXtrian(2) + Cote(2) * DXtrian(3)) * DemiPerMoinsCote(1) / Lloc0
' ******************************************************************
' *************     Centre du cercle circonscrit     ***************
' ************* point d'intersection des médiatrices ***************
' ******************************************************************
Xcir = (2 * Xmediane(1) * DXtrian(1) * DYtrian(2) + DYtrian(1) * DYtrian(2) * DYtrian(3) - 2 * Xmediane(2) * DXtrian(2) * DYtrian(1)) / 2 / lloc
Ycir = -(2 * Ymediane(1) * DYtrian(1) * DXtrian(2) + DXtrian(1) * DXtrian(2) * DXtrian(3) - 2 * Ymediane(2) * DYtrian(2) * DXtrian(1)) / 2 / lloc
' **************************************************************
' *************     Centre du cercle d'Euler     ***************
' *************       cercle aux 9 points        ***************
' **************************************************************
Xeul = (Xort + Xcir) / 2
Yeul = (Yort + Ycir) / 2
' **************************************************************
End Sub




Private Sub Form_Load()
   '***************************************************
   '***** mise en place de la grille des sommets ******
   '***************************************************
   FenetreTriangle.GridSom.FixedAlignment(0) = 2
   FenetreTriangle.GridSom.FixedAlignment(1) = 2
   FenetreTriangle.GridSom.FixedAlignment(2) = 2
   FenetreTriangle.GridSom.ColWidth(0) = 1000
   FenetreTriangle.GridSom.ColWidth(1) = 1700
   FenetreTriangle.GridSom.ColWidth(2) = 1700
   FenetreTriangle.GridSom.Row = 0
   FenetreTriangle.GridSom.Col = 0
   FenetreTriangle.GridSom.Text = "Sommets"
   FenetreTriangle.GridSom.Col = 1
   FenetreTriangle.GridSom.Text = "X"
   FenetreTriangle.GridSom.Col = 2
   FenetreTriangle.GridSom.Text = "Y"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreTriangle.GridSom.Col = 0
   FenetreTriangle.GridSom.Row = 1
   FenetreTriangle.GridSom.Text = "A"
   FenetreTriangle.GridSom.Row = 2
   FenetreTriangle.GridSom.Text = "B"
   FenetreTriangle.GridSom.Row = 3
   FenetreTriangle.GridSom.Text = "C"
   ' ********************************
   '     Valeurs par défaut
   ' ********************************
   Xtrian(1) = 4
   Ytrian(1) = 4
   Xtrian(2) = 0
   Ytrian(2) = 0
   Xtrian(3) = 6
   Ytrian(3) = 0
   ' *********************************************
   ' ***** affichages des valeurs par défaut *****
   ' *********************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000")
   Next i%
   '*****************************************************
   '** mise en place de la grille des angles et côtés ***
   '*****************************************************
   FenetreTriangle.GridAngCot.FixedAlignment(0) = 2
   FenetreTriangle.GridAngCot.FixedAlignment(1) = 2
   FenetreTriangle.GridAngCot.FixedAlignment(2) = 2
   FenetreTriangle.GridAngCot.ColWidth(0) = 1000
   FenetreTriangle.GridAngCot.ColWidth(1) = 1700
   FenetreTriangle.GridAngCot.ColWidth(2) = 1700
   ' ********************************
   ' libellés 1ère ligne
   ' ********************************
   FenetreTriangle.GridAngCot.Row = 0
   'FenetreTriangle.gridAngCot.Col = 0
   'FenetreTriangle.gridAngCot.Text = "Angles et côtés"
   FenetreTriangle.GridAngCot.Col = 1
   FenetreTriangle.GridAngCot.Text = "Angle (°)"
   FenetreTriangle.GridAngCot.Col = 2
   FenetreTriangle.GridAngCot.Text = "Côté opposé"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetreTriangle.GridAngCot.Col = 0
   FenetreTriangle.GridAngCot.Row = 1
   FenetreTriangle.GridAngCot.Text = "A et a"
   FenetreTriangle.GridAngCot.Row = 2
   FenetreTriangle.GridAngCot.Text = "B et b"
   FenetreTriangle.GridAngCot.Row = 3
   FenetreTriangle.GridAngCot.Text = "C et c"
   '******************************************************************
   '***      mise en place de la grille des             **************
   '*** médianes, hauteurs,bissectrices et médiatrices  **************              ******
   '******************************************************************
   FenetreTriangle.GridMHBM.FixedAlignment(0) = 2
   FenetreTriangle.GridMHBM.FixedAlignment(1) = 2
   FenetreTriangle.GridMHBM.FixedAlignment(2) = 2
   FenetreTriangle.GridMHBM.FixedAlignment(3) = 2
   FenetreTriangle.GridMHBM.ColWidth(0) = 1400
   FenetreTriangle.GridMHBM.ColWidth(1) = 800
   FenetreTriangle.GridMHBM.ColWidth(2) = 800
   FenetreTriangle.GridMHBM.ColWidth(3) = 800
   ' ********************************
   ' libellés 1ère ligne
   ' ********************************
   FenetreTriangle.GridMHBM.Row = 0
   FenetreTriangle.GridMHBM.Col = 1
   FenetreTriangle.GridMHBM.Text = "A"
   FenetreTriangle.GridMHBM.Col = 2
   FenetreTriangle.GridMHBM.Text = "B"
   FenetreTriangle.GridMHBM.Col = 3
   FenetreTriangle.GridMHBM.Text = "C"
   ' ********************************
   ' libellés 1ère colonne
   ' ********************************
   FenetreTriangle.GridMHBM.Col = 0
   FenetreTriangle.GridMHBM.Row = 0
   FenetreTriangle.GridMHBM.Text = "longueurs :"
   FenetreTriangle.GridMHBM.Row = 1
   FenetreTriangle.GridMHBM.Text = "Médianes"
   FenetreTriangle.GridMHBM.Row = 2
   FenetreTriangle.GridMHBM.Text = "Hauteurs"
   FenetreTriangle.GridMHBM.Row = 3
   FenetreTriangle.GridMHBM.Text = "Bissectrices"
   FenetreTriangle.GridMHBM.Row = 4
   FenetreTriangle.GridMHBM.Text = "Médiatrices"
   '*****************************************************
   '       mise en place de la grille des    ************
   '         Barycentre, Orthocentre,        ************
   ' cercles Inscrit, Circonscrit et d'Euler ************
   '*****************************************************
   FenetreTriangle.GridCentres.FixedAlignment(0) = 2
   FenetreTriangle.GridCentres.FixedAlignment(1) = 2
   FenetreTriangle.GridCentres.FixedAlignment(2) = 2
   FenetreTriangle.GridCentres.FixedAlignment(3) = 2
   FenetreTriangle.GridCentres.ColWidth(0) = 1600
   FenetreTriangle.GridCentres.ColWidth(1) = 800
   FenetreTriangle.GridCentres.ColWidth(2) = 800
   FenetreTriangle.GridCentres.ColWidth(3) = 800
   ' ********************************
   ' libellés 1ère ligne
   ' ********************************
   FenetreTriangle.GridCentres.Row = 0
   FenetreTriangle.GridCentres.Col = 1
   FenetreTriangle.GridCentres.Text = "X"
   FenetreTriangle.GridCentres.Col = 2
   FenetreTriangle.GridCentres.Text = "Y"
   FenetreTriangle.GridCentres.Col = 3
   FenetreTriangle.GridCentres.Text = "rayon"
   ' ********************************
   ' libellés 1ère colonne
   ' ********************************
   FenetreTriangle.GridCentres.Col = 0
   FenetreTriangle.GridCentres.Row = 0
   FenetreTriangle.GridCentres.Text = "positions :"
   FenetreTriangle.GridCentres.Row = 1
   FenetreTriangle.GridCentres.Text = "Barycentre"
   FenetreTriangle.GridCentres.Row = 2
   FenetreTriangle.GridCentres.Text = "Orthocentre"
   FenetreTriangle.GridCentres.Row = 3
   FenetreTriangle.GridCentres.Text = "Cercle inscrit"
   FenetreTriangle.GridCentres.Row = 4
   FenetreTriangle.GridCentres.Text = "Cercle circonscrit"
   FenetreTriangle.GridCentres.Row = 5
   FenetreTriangle.GridCentres.Text = "Cercle d'Euler"
   ' ***********************************
   ' Drapeau indiquant si CalculsDivers
   ' a été effectué ou non
   ' ***********************************
   TrianFlag = False
   ' ***********************************
   ' Drapeau d'erreur
   ' ***********************************
   ErrTrian = False
   ' ****************************************************
End Sub

Private Sub gridAngCot_KeyPress(KeyAscii As Integer)
   If lblCaracTriangle.Caption = "Données : côtés et angles" Then
      If FenetreTriangle.GridAngCot.Row = 1 And FenetreTriangle.GridAngCot.Col = 1 Then
         Beep
         Exit Sub
      ElseIf FenetreTriangle.GridAngCot.Row = 2 And FenetreTriangle.GridAngCot.Col = 2 Then
         Beep
         Exit Sub
      ElseIf FenetreTriangle.GridAngCot.Row = 3 And FenetreTriangle.GridAngCot.Col = 1 Then
         Beep
         Exit Sub
      End If
      AngCotText$ = GridAngCot.Text
      Select Case KeyAscii
      Case 32 To 168
         AngCotCar$ = Chr(KeyAscii)
         AngCotText$ = AngCotText$ & AngCotCar$
         GridAngCot.Text = AngCotText$
      Case 8
         If Len(GridAngCot.Text) > 0 Then
            AngCotText$ = Left$(AngCotText$, Len(AngCotText$) - 1)
            GridAngCot.Text = AngCotText$
         Else
            Beep
         End If
      End Select
   End If
End Sub


Private Sub gridSom_KeyPress(KeyAscii As Integer)
   If lblCaracTriangle.Caption = "Données : coordonnées des sommets" Then
      SomText$ = GridSom.Text
      Select Case KeyAscii
      Case 32 To 168
         SomCar$ = Chr(KeyAscii)
         SomText$ = SomText$ & SomCar$
         GridSom.Text = SomText$
      Case 8
         If Len(GridSom.Text) > 0 Then
            SomText$ = Left$(SomText$, Len(SomText$) - 1)
            GridSom.Text = SomText$
         Else
            Beep
         End If
      End Select
   End If
End Sub



Private Sub mnuCalculer_Click()
   ErrTrian = False
   If lblCaracTriangle.Caption = "Données : côtés et angles" Then
      GridAngCot.Row = 1
      GridAngCot.Col = 2
      Cote(1) = CSng(GridAngCot.Text)
      GridAngCot.Row = 2
      GridAngCot.Col = 1
      AngleTrian(2) = CSng(GridAngCot.Text) / CONVDEGRAD
      GridAngCot.Row = 3
      GridAngCot.Col = 2
      Cote(3) = CSng(GridAngCot.Text)
      Call AngCotVersSom
   Else
      For i% = 1 To 3
         GridSom.Row = i%
         GridSom.Col = 1
         Xtrian(i%) = CSng(GridSom.Text)
         GridSom.Col = 2
         Ytrian(i%) = CSng(GridSom.Text)
      Next i%
      Call SomVersAngCot
   End If
   If ErrTrian = True Then
      ErrTrian = False
      Exit Sub
   End If
   Call CalculsDivers
   TrianFlag = True
   Call PresenteResultats
End Sub





Private Sub mnuDonneesCoor_Click()
   ' données : coordonnées des sommets
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle à définir :"
   lblCaracTriangle.Caption = "Données : coordonnées des sommets"
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub


Private Sub mnuDonneesCotAng_Click()
   ' données : angles et côtés
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle à définir :"
   lblCaracTriangle.Caption = "Données : côtés et angles"
   ' **************************************
   ' **** remise à zéro grille sommets ****
   ' **************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = ""
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = ""
   Next i%
   ' *****************************************
   ' **** interdiction de certaines cases ****
   ' *****************************************
   FenetreTriangle.GridAngCot.Row = 1
   FenetreTriangle.GridAngCot.Col = 1
   FenetreTriangle.GridAngCot.Text = "XXXXX"
   FenetreTriangle.GridAngCot.Row = 2
   FenetreTriangle.GridAngCot.Col = 2
   FenetreTriangle.GridAngCot.Text = "XXXXX"
   FenetreTriangle.GridAngCot.Row = 3
   FenetreTriangle.GridAngCot.Col = 1
   FenetreTriangle.GridAngCot.Text = "XXXXX"
   ' ******************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub

Private Sub mnuEquilateral_Click()
   ' triangle équilatéral
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle exemple :"
   lblCaracTriangle.Caption = "Equilatéral"
   Xtrian(1) = 0
   Ytrian(1) = 1.732051
   Xtrian(2) = -1
   Ytrian(2) = 0
   Xtrian(3) = 1
   Ytrian(3) = 0
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000000")
   Next i%
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub

Private Sub mnuIsocele_Click()
   ' triangle isocèle
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle exemple :"
   lblCaracTriangle.Caption = "Isocèle"
   Xtrian(1) = 0
   Ytrian(1) = 5
   Xtrian(2) = -2
   Ytrian(2) = 0
   Xtrian(3) = 2
   Ytrian(3) = 0
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000000")
   Next i%
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub

Private Sub mnuQuitter_Click()
   FenetreTriangle.Hide
End Sub




Private Sub mnuRectangle_Click()
   ' triangle rectangle
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle exemple :"
   lblCaracTriangle.Caption = "Isocèle"
   Xtrian(1) = 0
   Ytrian(1) = 2
   Xtrian(2) = 0
   Ytrian(2) = 0
   Xtrian(3) = 1
   Ytrian(3) = 0
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000000")
   Next i%
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub


Private Sub mnuRectIso_Click()
   ' triangle rectangle isocèle
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle exemple :"
   lblCaracTriangle.Caption = "Rectangle isocèle"
   Xtrian(1) = 0
   Ytrian(1) = 1
   Xtrian(2) = -1
   Ytrian(2) = 0
   Xtrian(3) = 1
   Ytrian(3) = 0
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000000")
   Next i%
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub

Private Sub mnuTroisAnglesAigus_Click()
   ' 3 angles aigus
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle exemple :"
   lblCaracTriangle.Caption = "Quelconque avec trois angles aigus"
   Xtrian(1) = 4
   Ytrian(1) = 4
   Xtrian(2) = 0
   Ytrian(2) = 0
   Xtrian(3) = 6
   Ytrian(3) = 0
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000000")
   Next i%
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub


Private Sub mnuUnAngleObtus_Click()
   ' 1 angle obtus
   TrianFlag = False
   lblTypeTriangle.Caption = "Triangle exemple :"
   lblCaracTriangle.Caption = "Quelconque avec un angle obtus"
   Xtrian(1) = 4
   Ytrian(1) = 2
   Xtrian(2) = 0
   Ytrian(2) = 0
   Xtrian(3) = 6
   Ytrian(3) = 0
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To 3
      FenetreTriangle.GridSom.Row = i%
      FenetreTriangle.GridSom.Col = 1
      FenetreTriangle.GridSom.Text = Format(Xtrian(i%), "0.000000")
      FenetreTriangle.GridSom.Col = 2
      FenetreTriangle.GridSom.Text = Format(Ytrian(i%), "0.000000")
   Next i%
   ' **********************************************
   ' **** remise à zéro grille angles et côtés ****
   ' **********************************************
   For i% = 1 To 3
      GridAngCot.Row = i%
      GridAngCot.Col = 1
      GridAngCot.Text = ""
      GridAngCot.Col = 2
      GridAngCot.Text = ""
   Next i%
   ' ***************************************************
   ' **************************************
   ' **** remise à zéro grille MHBM ****
   ' **************************************
   For i% = 1 To 4
      FenetreTriangle.GridMHBM.Row = i%
      FenetreTriangle.GridMHBM.Col = 1
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 2
      FenetreTriangle.GridMHBM.Text = ""
      FenetreTriangle.GridMHBM.Col = 3
      FenetreTriangle.GridMHBM.Text = ""
   Next i%
   ' **************************************
   ' **** remise à zéro grille Centres ****
   ' **************************************
   For i% = 1 To 5
      FenetreTriangle.GridCentres.Row = i%
      FenetreTriangle.GridCentres.Col = 1
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 2
      FenetreTriangle.GridCentres.Text = ""
      FenetreTriangle.GridCentres.Col = 3
      FenetreTriangle.GridCentres.Text = ""
   Next i%
   ' *********************************************
   ' **** remise à zéro étiquette lblValDemiP ****
   ' *********************************************
   FenetreTriangle.lblValDemiP.Caption = ""
   ' ********************************************
   ' **** remise à zéro étiquette lblValSurf ****
   ' ********************************************
   FenetreTriangle.lblValSurf.Caption = ""
   ' ********************************************
End Sub

Private Sub mnuVisualiser_Click()
   If TrianFlag = False Then
      mnuCalculer_Click
   End If
   FenetreVisuTriangle.Show
End Sub


Private Sub PresenteResultats()
   '****************************************************
   '            Demi-périmètre et Surface
   '****************************************************
   lblValDemiP.Caption = Format(DemiPerTrian, "0.000")
   lblValSurf.Caption = Format(AireTrian, "0.000")
   '****************************************************
   'grille médianes, hauteurs, bissectrices, médiatrices
   '****************************************************
   GridMHBM.Row = 1
   For i% = 1 To 3
      GridMHBM.Col = i%
      GridMHBM.Text = Format(Lmediane(i%), "0.000")
   Next i%
   GridMHBM.Row = 2
   For i% = 1 To 3
      GridMHBM.Col = i%
      GridMHBM.Text = Format(Lhauteur(i%), "0.000")
   Next i%
   GridMHBM.Row = 3
   For i% = 1 To 3
      GridMHBM.Col = i%
      GridMHBM.Text = Format(Lbissec(i%), "0.000")
   Next i%
   GridMHBM.Row = 4
   For i% = 1 To 3
      GridMHBM.Col = i%
      GridMHBM.Text = Format(Lmediat(i%), "0.000")
   Next i%
   '****************************************************
   '*****      grille Barycentre, Orthocentre,    ******
   '*****  Cercle inscrit, circonscrit et d'Euler ******
   '****************************************************
   GridCentres.Row = 1
   GridCentres.Col = 1
   GridCentres.Text = Format(Xbar, "0.000")
   GridCentres.Col = 2
   GridCentres.Text = Format(Ybar, "0.000")
   '+++++++++++++++++++++++++++++++++++++++
   GridCentres.Row = 2
   GridCentres.Col = 1
   GridCentres.Text = Format(Xort, "0.000")
   GridCentres.Col = 2
   GridCentres.Text = Format(Yort, "0.000")
   '+++++++++++++++++++++++++++++++++++++++
   GridCentres.Row = 3
   GridCentres.Col = 1
   GridCentres.Text = Format(Xins, "0.000")
   GridCentres.Col = 2
   GridCentres.Text = Format(Yins, "0.000")
   GridCentres.Col = 3
   GridCentres.Text = Format(RayonInscrit, "0.000")
   '+++++++++++++++++++++++++++++++++++++++
   GridCentres.Row = 4
   GridCentres.Col = 1
   GridCentres.Text = Format(Xcir, "0.000")
   GridCentres.Col = 2
   GridCentres.Text = Format(Ycir, "0.000")
   GridCentres.Col = 3
   GridCentres.Text = Format(RayonCirc, "0.000")
   '+++++++++++++++++++++++++++++++++++++++
   GridCentres.Row = 5
   GridCentres.Col = 1
   GridCentres.Text = Format(Xeul, "0.000")
   GridCentres.Col = 2
   GridCentres.Text = Format(Yeul, "0.000")
   GridCentres.Col = 3
   GridCentres.Text = Format(RayonEuler, "0.000")
   '+++++++++++++++++++++++++++++++++++++++
End Sub

Private Sub SomVersAngCot()
' ----------------------------------------------------
' Calcule les angles et les côtés à partir des sommets
' ----------------------------------------------------
' coordonnées des sommets
'Public Xtrian(1 To 3), Ytrian(1 To 3)
' différences de coordonnées des sommets
'Public DXtrian(1 To 3), DYtrian(1 To 3)
' angles et sommets
'Public AngleTrian(1 To 3), Cote(1 To 3)
'Public Const CONVDEGRAD = 180/PI
' ----------------------------------------------------
For i% = 1 To 3
   j% = i% + 1
   If j% = 4 Then j% = 1
   k% = j% + 1
   If k% = 4 Then k% = 1
   DXtrian(i%) = Xtrian(k%) - Xtrian(j%)
   DYtrian(i%) = Ytrian(k%) - Ytrian(j%)
   Cote(i%) = Sqr(DXtrian(i%) * DXtrian(i%) + DYtrian(i%) * DYtrian(i%))
Next i%
' demi-périmètre :
DemiPerTrian = (Cote(1) + Cote(2) + Cote(3)) / 2
If DemiPerTrian <= 0 Then
   ErrTrian = True
   Beep
   MsgBox "Erreur dans les données !", 48, "TRIANGLE"
   Exit Sub
End If
' aire
AireTrian = 0
' rayon du cercle circonscrit (médiatrices)
RayonCirc = 1
For i% = 1 To 3
   DemiPerMoinsCote(i%) = DemiPerTrian - Cote(i%)
   AireTrian = AireTrian - Xtrian(i%) * DYtrian(i%)
   RayonCirc = RayonCirc * Cote(i%)
Next i%
AireTrian = AireTrian / 2
If AireTrian <= 0 Then
   AireTrian = -AireTrian
   ' A, B et C ne se succèdent pas
   ' dans le sens trigonométrique
   MsgBox "A, B et C ne se succèdent pas dans le sens trigonométrique", 64, "TRIANGLE"
End If
RayonInscrit = AireTrian / DemiPerTrian
RayonCirc = RayonCirc / AireTrian / 4
RayonEuler = RayonCirc / 2
For i% = 1 To 3
   TangDemiAng = RayonInscrit / DemiPerMoinsCote(i%)
   AngleTrian(i%) = 2 * Atn(TangDemiAng)
   TanP = 1 - TangDemiAng * TangDemiAng
   TanM = 1 + TangDemiAng * TangDemiAng
   SinAng(i%) = 2 * TangDemiAng / TanP
   CosAng(i%) = TanM / TanP
Next i%
' ***************************************************
'           Affichage angles et côtés
' ***************************************************
For i% = 1 To 3
   GridAngCot.Row = i%
   GridAngCot.Col = 1
   GridAngCot.Text = Format(AngleTrian(i%) * CONVDEGRAD, "000.00")
   GridAngCot.Col = 2
   GridAngCot.Text = Format(Cote(i%), "0.000")
Next i%
' ***************************************************
End Sub


