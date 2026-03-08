VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetrePolyedre 
   Caption         =   "Polyèdres"
   ClientHeight    =   5790
   ClientLeft      =   975
   ClientTop       =   975
   ClientWidth     =   7785
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
   ScaleHeight     =   5790
   ScaleWidth      =   7785
   Begin MSFlexGridLib.MSFlexGrid GridSom 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6588
      _Version        =   393216
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
   Begin VB.TextBox txtNbSom 
      Alignment       =   2  'Center
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
      Left            =   2400
      TabIndex        =   1
      Text            =   "4"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblNomPol 
      Caption         =   "Tétraèdre"
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
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label lblTypePol 
      Caption         =   "Polyèdre régulier :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblNbSom 
      Caption         =   "Nombre de sommets :"
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
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblCoorPts 
      Caption         =   "Coordonnées des sommets :"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu mnuEnregistrer 
         Caption         =   "&Enregistrer..."
      End
   End
   Begin VB.Menu mnuPolyedre 
      Caption         =   "&Polyèdre"
      Begin VB.Menu mnuRegulier 
         Caption         =   "&Régulier"
         Begin VB.Menu mnuTetraedre 
            Caption         =   "&Tétraèdre"
         End
         Begin VB.Menu mnuOctaedre 
            Caption         =   "&Octaèdre"
         End
         Begin VB.Menu mnuCube 
            Caption         =   "&Cube"
         End
         Begin VB.Menu mnuIcosaedre 
            Caption         =   "&Icosaèdre"
         End
         Begin VB.Menu mnuPentagonoDodecaedre 
            Caption         =   "&Pentagono-dodécaèdre"
         End
      End
      Begin VB.Menu mnuParticulier 
         Caption         =   "&Particulier"
         Begin VB.Menu mnuCuboctaedre1 
            Caption         =   "Cuboctaedre 1"
         End
         Begin VB.Menu mnuCuboctaedre2 
            Caption         =   "Cuboctaedre 2"
         End
         Begin VB.Menu mnuPentagonoDoIco 
            Caption         =   "Pentagono-dodécaèdre icosaedre"
         End
         Begin VB.Menu mnuBallFoot 
            Caption         =   """Ballon de &football"""
         End
         Begin VB.Menu mnuRhombododecaedre 
            Caption         =   "Rhombododécaèdre"
         End
         Begin VB.Menu mnuTetragonoTriOctaedre 
            Caption         =   "Tétragono-tri-octaèdre"
         End
      End
      Begin VB.Menu mnuDefiniUt 
         Caption         =   "&Défini par l'utilisateur"
      End
   End
   Begin VB.Menu mnuVisu 
      Caption         =   "&Visualiser"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
   Begin VB.Menu mnuAide 
      Caption         =   "&?"
   End
End
Attribute VB_Name = "FenetrePolyedre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Bascule :
Dim DefiniUt As Boolean
' FALSE si le polyèdre choisi est pré-défini,
' TRUE s'il est défini par l'utilisateur.



Private Sub ballondefoot()
   NbSom% = 60     ' Nombre de sommets
   NbFacMax% = 116     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = UAT
   Ysom(1) = 0
   Zsom(1) = -3
   Xsom(2) = 2 * UAT
   Ysom(2) = 1
   Zsom(2) = -TAU - 1
   Xsom(3) = UAT
   Ysom(3) = 2
   Zsom(3) = -2 * TAU + 1
   Xsom(4) = -UAT
   Ysom(4) = 2
   Zsom(4) = -2 * TAU + 1
   Xsom(5) = -2 * UAT
   Ysom(5) = 1
   Zsom(5) = -TAU - 1
   Xsom(6) = -UAT
   Ysom(6) = 0
   Zsom(6) = -3
   Xsom(7) = -2 * UAT
   Ysom(7) = -1
   Zsom(7) = -TAU - 1
   Xsom(8) = -UAT
   Ysom(8) = -2
   Zsom(8) = -2 * TAU + 1
   Xsom(9) = UAT
   Ysom(9) = -2
   Zsom(9) = -2 * TAU + 1
   Xsom(10) = 2 * UAT
   Ysom(10) = -1
   Zsom(10) = -TAU - 1
   Xsom(11) = 2 * TAU - 1
   Ysom(11) = -UAT
   Zsom(11) = -2
   Xsom(12) = 2 * TAU - 1
   Ysom(12) = UAT
   Zsom(12) = -2
   Xsom(13) = TAU + 1
   Ysom(13) = 2 * UAT
   Zsom(13) = -1
   Xsom(14) = 2
   Ysom(14) = 2 * TAU - 1
   Zsom(14) = -UAT
   Xsom(15) = 1
   Ysom(15) = TAU + 1
   Zsom(15) = -2 * UAT
   Xsom(16) = 0
   Ysom(16) = 3
   Zsom(16) = -UAT
   Xsom(17) = -1
   Ysom(17) = TAU + 1
   Zsom(17) = -2 * UAT
   Xsom(18) = -2
   Ysom(18) = 2 * TAU - 1
   Zsom(18) = -UAT
   Xsom(19) = -TAU - 1
   Ysom(19) = 2 * UAT
   Zsom(19) = -1
   Xsom(20) = -2 * TAU + 1
   Ysom(20) = UAT
   Zsom(20) = -2
   Xsom(21) = -2 * TAU + 1
   Ysom(21) = -UAT
   Zsom(21) = -2
   Xsom(22) = -TAU - 1
   Ysom(22) = -2 * UAT
   Zsom(22) = -1
   Xsom(23) = -2
   Ysom(23) = -2 * TAU + 1
   Zsom(23) = -UAT
   Xsom(24) = -1
   Ysom(24) = -TAU - 1
   Zsom(24) = -2 * UAT
   Xsom(25) = 0
   Ysom(25) = -3
   Zsom(25) = -UAT
   Xsom(26) = 1
   Ysom(26) = -TAU - 1
   Zsom(26) = -2 * UAT
   Xsom(27) = 2
   Ysom(27) = -2 * TAU + 1
   Zsom(27) = -UAT
   Xsom(28) = TAU + 1
   Ysom(28) = -2 * UAT
   Zsom(28) = -1
   Xsom(29) = 3
   Ysom(29) = -UAT
   Zsom(29) = 0
   Xsom(30) = 3
   Ysom(30) = UAT
   Zsom(30) = 0
   Xsom(31) = -3
   Ysom(31) = UAT
   Zsom(31) = 0
   Xsom(32) = -3
   Ysom(32) = -UAT
   Zsom(32) = 0
   Xsom(33) = -TAU - 1
   Ysom(33) = -2 * UAT
   Zsom(33) = 1
   Xsom(34) = -2
   Ysom(34) = -2 * TAU + 1
   Zsom(34) = UAT
   Xsom(35) = -1
   Ysom(35) = -TAU - 1
   Zsom(35) = 2 * UAT
   Xsom(36) = 0
   Ysom(36) = -3
   Zsom(36) = UAT
   Xsom(37) = 1
   Ysom(37) = -TAU - 1
   Zsom(37) = 2 * UAT
   Xsom(38) = 2
   Ysom(38) = -2 * TAU + 1
   Zsom(38) = UAT
   Xsom(39) = TAU + 1
   Ysom(39) = -2 * UAT
   Zsom(39) = 1
   Xsom(40) = 2 * TAU - 1
   Ysom(40) = -UAT
   Zsom(40) = 2
   Xsom(41) = 2 * TAU - 1
   Ysom(41) = UAT
   Zsom(41) = 2
   Xsom(42) = TAU + 1
   Ysom(42) = 2 * UAT
   Zsom(42) = 1
   Xsom(43) = 2
   Ysom(43) = 2 * TAU - 1
   Zsom(43) = UAT
   Xsom(44) = 1
   Ysom(44) = TAU + 1
   Zsom(44) = 2 * UAT
   Xsom(45) = 0
   Ysom(45) = 3
   Zsom(45) = UAT
   Xsom(46) = -1
   Ysom(46) = TAU + 1
   Zsom(46) = 2 * UAT
   Xsom(47) = -2
   Ysom(47) = 2 * TAU - 1
   Zsom(47) = UAT
   Xsom(48) = -TAU - 1
   Ysom(48) = 2 * UAT
   Zsom(48) = 1
   Xsom(49) = -2 * TAU + 1
   Ysom(49) = UAT
   Zsom(49) = 2
   Xsom(50) = -2 * TAU + 1
   Ysom(50) = -UAT
   Zsom(50) = 2
   Xsom(51) = -2 * UAT
   Ysom(51) = -1
   Zsom(51) = TAU + 1
   Xsom(52) = -UAT
   Ysom(52) = -2
   Zsom(52) = 2 * TAU - 1
   Xsom(53) = UAT
   Ysom(53) = -2
   Zsom(53) = 2 * TAU - 1
   Xsom(54) = 2 * UAT
   Ysom(54) = -1
   Zsom(54) = TAU + 1
   Xsom(55) = UAT
   Ysom(55) = 0
   Zsom(55) = 3
   Xsom(56) = 2 * UAT
   Ysom(56) = 1
   Zsom(56) = TAU + 1
   Xsom(57) = UAT
   Ysom(57) = 2
   Zsom(57) = 2 * TAU - 1
   Xsom(58) = -UAT
   Ysom(58) = 2
   Zsom(58) = 2 * TAU - 1
   Xsom(59) = -2 * UAT
   Ysom(59) = 1
   Zsom(59) = TAU + 1
   Xsom(60) = -UAT
   Ysom(60) = 0
   Zsom(60) = 3
   ' ****************************************************
End Sub





Private Sub cube()
   NbSom% = 8     ' Nombre de sommets
   NbFacMax% = 12     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 1
   Ysom(1) = 1
   Zsom(1) = 1
   Xsom(2) = -1
   Ysom(2) = 1
   Zsom(2) = 1
   Xsom(3) = -1
   Ysom(3) = -1
   Zsom(3) = 1
   Xsom(4) = 1
   Ysom(4) = -1
   Zsom(4) = 1
   Xsom(5) = 1
   Ysom(5) = 1
   Zsom(5) = -1
   Xsom(6) = -1
   Ysom(6) = 1
   Zsom(6) = -1
   Xsom(7) = -1
   Ysom(7) = -1
   Zsom(7) = -1
   Xsom(8) = 1
   Ysom(8) = -1
   Zsom(8) = -1
   ' ****************************************************
   ' *****************************************************
   ' ***************  valeurs des arêtes : ***************
   ' *****************************************************
   Arete%(1, 2) = 1
   Arete%(1, 3) = 0
   Arete%(1, 4) = 1
   Arete%(1, 5) = 1
   Arete%(1, 6) = 0
   Arete%(1, 7) = 0
   Arete%(1, 8) = 0
   Arete%(2, 3) = 1
   Arete%(2, 4) = 0
   Arete%(2, 5) = 0
   Arete%(2, 6) = 1
   Arete%(2, 7) = 0
   Arete%(2, 8) = 0
   Arete%(3, 4) = 1
   Arete%(3, 5) = 0
   Arete%(3, 6) = 0
   Arete%(3, 7) = 1
   Arete%(3, 8) = 0
   Arete%(4, 5) = 0
   Arete%(4, 6) = 0
   Arete%(4, 7) = 0
   Arete%(4, 8) = 1
   Arete%(5, 6) = 2
   Arete%(5, 7) = 0
   Arete%(5, 8) = 2
   Arete%(6, 7) = 2
   Arete%(6, 8) = 0
   Arete%(7, 8) = 2
   ' *****************************************************
   RayonMax = Sqr(3)
   ' *****************************************************
End Sub

Private Sub cuboctaedre1()
   NbSom% = 12     ' Nombre de sommets
   NbFacMax% = 24     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 1
   Ysom(1) = 0
   Zsom(1) = -1
   Xsom(2) = 0
   Ysom(2) = 1
   Zsom(2) = -1
   Xsom(3) = -1
   Ysom(3) = 0
   Zsom(3) = -1
   Xsom(4) = 0
   Ysom(4) = -1
   Zsom(4) = -1
   Xsom(5) = 1
   Ysom(5) = 1
   Zsom(5) = 0
   Xsom(6) = -1
   Ysom(6) = 1
   Zsom(6) = 0
   Xsom(7) = -1
   Ysom(7) = -1
   Zsom(7) = 0
   Xsom(8) = 1
   Ysom(8) = -1
   Zsom(8) = 0
   Xsom(9) = 1
   Ysom(9) = 0
   Zsom(9) = 1
   Xsom(10) = 0
   Ysom(10) = 1
   Zsom(10) = 1
   Xsom(11) = -1
   Ysom(11) = 0
   Zsom(11) = 1
   Xsom(12) = 0
   Ysom(12) = -1
   Zsom(12) = 1
   ' ****************************************************
End Sub

Private Sub cuboctaedre2()
   NbSom% = 24     ' Nombre de sommets
   NbFacMax% = 48     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 1
   Ysom(1) = 0
   Zsom(1) = -2
   Xsom(2) = 0
   Ysom(2) = 1
   Zsom(2) = -2
   Xsom(3) = -1
   Ysom(3) = 0
   Zsom(3) = -2
   Xsom(4) = 0
   Ysom(4) = -1
   Zsom(4) = -2
   Xsom(5) = 2
   Ysom(5) = 0
   Zsom(5) = -1
   Xsom(6) = 0
   Ysom(6) = 2
   Zsom(6) = -1
   Xsom(7) = -2
   Ysom(7) = 0
   Zsom(7) = -1
   Xsom(8) = 0
   Ysom(8) = -2
   Zsom(8) = -1
   Xsom(9) = 2
   Ysom(9) = 1
   Zsom(9) = 0
   Xsom(10) = 1
   Ysom(10) = 2
   Zsom(10) = 0
   Xsom(11) = -1
   Ysom(11) = 2
   Zsom(11) = 0
   Xsom(12) = -2
   Ysom(12) = 1
   Zsom(12) = 0
   Xsom(13) = -2
   Ysom(13) = -1
   Zsom(13) = 0
   Xsom(14) = -1
   Ysom(14) = -2
   Zsom(14) = 0
   Xsom(15) = 1
   Ysom(15) = -2
   Zsom(15) = 0
   Xsom(16) = 2
   Ysom(16) = -1
   Zsom(16) = 0
   Xsom(17) = 2
   Ysom(17) = 0
   Zsom(17) = 1
   Xsom(18) = 0
   Ysom(18) = 2
   Zsom(18) = 1
   Xsom(19) = -2
   Ysom(19) = 0
   Zsom(19) = 1
   Xsom(20) = 0
   Ysom(20) = -2
   Zsom(20) = 1
   Xsom(21) = 1
   Ysom(21) = 0
   Zsom(21) = 2
   Xsom(22) = 0
   Ysom(22) = 1
   Zsom(22) = 2
   Xsom(23) = -1
   Ysom(23) = 0
   Zsom(23) = 2
   Xsom(24) = 0
   Ysom(24) = -1
   Zsom(24) = 2
   ' ****************************************************
End Sub

Private Sub Form_Load()
   mnuOuvrir.Enabled = False
   DefiniUt = False
   NomPolyedre$ = "Tétraèdre"
   Call tetraedre
   '***************************************
   '***** mise en place de la grille ******
   '***************************************
   txtNbSom.Text = CStr(NbSom%)
   GridSom.Rows = NbSom% + 1
   FenetrePolyedre.GridSom.FixedAlignment(0) = 2
   FenetrePolyedre.GridSom.FixedAlignment(1) = 2
   FenetrePolyedre.GridSom.FixedAlignment(2) = 2
   FenetrePolyedre.GridSom.FixedAlignment(3) = 2
   FenetrePolyedre.GridSom.ColWidth(1) = 2000
   FenetrePolyedre.GridSom.ColWidth(2) = 2000
   FenetrePolyedre.GridSom.ColWidth(3) = 2000
   FenetrePolyedre.GridSom.Row = 0
   FenetrePolyedre.GridSom.Col = 1
   FenetrePolyedre.GridSom.Text = "X"
   FenetrePolyedre.GridSom.Col = 2
   FenetrePolyedre.GridSom.Text = "Y"
   FenetrePolyedre.GridSom.Col = 3
   FenetrePolyedre.GridSom.Text = "Z"
   ' ********************************
   ' numérotation 1ère colonne
   ' ********************************
   FenetrePolyedre.GridSom.Col = 0
   For i% = 1 To NbSom%
      FenetrePolyedre.GridSom.Row = i%
      FenetrePolyedre.GridSom.Text = Format(i%, "0")
   Next i%
   ' ********************************
   ' alignement des nombres à droite
   ' ********************************
   For i% = 1 To 3
      FenetrePolyedre.GridSom.ColAlignment(i%) = 1
   Next i%
   ' ********************************
   ' placement des valeurs par défaut
   ' ********************************
   For i% = 1 To NbSom%
      FenetrePolyedre.GridSom.Row = i%
      FenetrePolyedre.GridSom.Col = 1
      FenetrePolyedre.GridSom.Text = Format(Xsom(i%), "0.000000000")
      FenetrePolyedre.GridSom.Col = 2
      FenetrePolyedre.GridSom.Text = Format(Ysom(i%), "0.000000000")
      FenetrePolyedre.GridSom.Col = 3
      FenetrePolyedre.GridSom.Text = Format(Zsom(i%), "0.000000000")
   Next i%
   ' **************************************
End Sub

Private Sub gridSom_KeyPress(KeyAscii As Integer)
   If DefiniUt = True Then ' si polyèdre quelconque
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

Private Sub icosaedre()
   NbSom% = 12     ' Nombre de sommets
   NbFacMax% = 20     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 1
   Ysom(1) = 0
   Zsom(1) = TAU
   Xsom(2) = -1
   Ysom(2) = 0
   Zsom(2) = TAU
   Xsom(3) = 0
   Ysom(3) = TAU
   Zsom(3) = 1
   Xsom(4) = 0
   Ysom(4) = -TAU
   Zsom(4) = 1
   Xsom(5) = TAU
   Ysom(5) = 1
   Zsom(5) = 0
   Xsom(6) = -TAU
   Ysom(6) = 1
   Zsom(6) = 0
   Xsom(7) = -TAU
   Ysom(7) = -1
   Zsom(7) = 0
   Xsom(8) = TAU
   Ysom(8) = -1
   Zsom(8) = 0
   Xsom(9) = 0
   Ysom(9) = TAU
   Zsom(9) = -1
   Xsom(10) = 0
   Ysom(10) = -TAU
   Zsom(10) = -1
   Xsom(11) = 1
   Ysom(11) = 0
   Zsom(11) = -TAU
   Xsom(12) = -1
   Ysom(12) = 0
   Zsom(12) = -TAU
   ' ****************************************************
End Sub






Private Sub mnuBallFoot_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre particulier"
   NomPolyedre$ = "'Ballon de Football'"
   lblNomPol.Caption = NomPolyedre$
   Call ballondefoot
   Call ActuGrille
End Sub

Private Sub mnuCube_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre régulier"
   NomPolyedre$ = "Cube"
   lblNomPol.Caption = NomPolyedre$
   Call cube
   Call ActuGrille
End Sub


Private Sub mnuCuboctaedre1_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre particulier"
   NomPolyedre$ = "Cuboctaèdre 1"
   lblNomPol.Caption = NomPolyedre$
   Call cuboctaedre1
   Call ActuGrille
End Sub

Private Sub mnuCuboctaedre2_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre particulier"
   NomPolyedre$ = "Cuboctaèdre 2"
   lblNomPol.Caption = NomPolyedre$
   Call cuboctaedre2
   Call ActuGrille
End Sub



Private Sub mnuDefiniUt_Click()
   DefiniUt = True
   mnuOuvrir.Enabled = True
   lblTypePol.Caption = "Polyèdre défini par l'utilisateur"
   NomPolyedre$ = ""
   lblNomPol.Caption = NomPolyedre$
End Sub

Private Sub mnuEnregistrer_Click()
   On Error GoTo Traite_ErreursEnrPolyedre
   Maths.ctrlCMDialog.DefaultExt = "xyz"
   Maths.ctrlCMDialog.Filter = "Points (*.xyz)|*.xyz"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   ' ********************************************
   ' Redims
   ' ********************************************
   Call PolReDim
   ' ********************************************
   ' affectation de leurs valeurs aux coordonnées
   ' ********************************************
   For i% = 1 To NbSom%
      FenetrePolyedre.GridSom.Row = i%
      FenetrePolyedre.GridSom.Col = 1
      Xsom(i%) = CSng(FenetrePolyedre.GridSom.Text)
      FenetrePolyedre.GridSom.Col = 2
      Ysom(i%) = CSng(FenetrePolyedre.GridSom.Text)
      FenetrePolyedre.GridSom.Col = 3
      Zsom(i%) = CSng(FenetrePolyedre.GridSom.Text)
   Next i%
   ' ********************************************
   '-----------------------------------------------------
   ' Création du fichier de points 3D
   ' et écriture de leurs coordonnées dans le fichier
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbSom%
   For i% = 1 To NbSom%
      Write #1, Xsom(i%)
      Write #1, Ysom(i%)
      Write #1, Zsom(i%)
   Next i%
   Close #1
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnrPolyedre:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub




Private Sub mnuIcosaedre_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre régulier"
   NomPolyedre$ = "Icosaèdre"
   lblNomPol.Caption = NomPolyedre$
   Call icosaedre
   Call ActuGrille
End Sub

Private Sub mnuOctaedre_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre régulier"
   NomPolyedre$ = "Octaèdre"
   lblNomPol.Caption = NomPolyedre$
   Call octaedre
   Call ActuGrille
End Sub

Private Sub mnuOuvrir_Click()
   '-----------------------------------------------------
   On Error GoTo Traite_ErreursOuvPolyedre
   Maths.ctrlCMDialog.Filter = "Points (*.xyz)|*.xyz"
   ' nom de fichier et chemin doivent exister
   ' sinon apparait un message d'erreur spécifique
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   '-----------------------------------------------------
   ' Ouverture et lecture du fichier de points
   ' et écriture de ces points dans Xsom(i%)
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, NbSom%
   ' ********************************************
   '               ReDims
   ' ********************************************
   Call PolReDim
   ' *************************************************
   ' lecture des éléments de Xsom, Ysom et Zsom
   ' *************************************************
   For i% = 1 To NbSom%
      Input #1, Xsom(i%)
      Input #1, Ysom(i%)
      Input #1, Zsom(i%)
   Next i%
   Close #1
   ' *************************************************
   ' placement du nombre de sommets, ce qui provoque
   ' le redimentionnement des grilles
   ' *************************************************
   txtNbSom.Text = Format(NbSom%, "0")
   ' ********************************
   ' placement des nouvelles valeurs
   ' ********************************
   For i% = 1 To NbSom%
      FenetrePolyedre.GridSom.Row = i%
      FenetrePolyedre.GridSom.Col = 1
      FenetrePolyedre.GridSom.Text = Format(Xsom(i%), "0.000000000")
      FenetrePolyedre.GridSom.Col = 2
      FenetrePolyedre.GridSom.Text = Format(Ysom(i%), "0.000000000")
      FenetrePolyedre.GridSom.Col = 3
      FenetrePolyedre.GridSom.Text = Format(Zsom(i%), "0.000000000")
   Next i%
   ' ********************************
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursOuvPolyedre:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub


Private Sub mnuPentagonoDodecaedre_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre régulier"
   NomPolyedre$ = "Pentagono-dodécaèdre"
   lblNomPol.Caption = NomPolyedre$
   Call pentagonododecaedre
   Call ActuGrille
End Sub


Private Sub mnuPentagonoDoIco_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre particulier"
   NomPolyedre$ = "Pentagono-dodécaèdre icosaèdre"
   lblNomPol.Caption = NomPolyedre$
   Call pentagonododecaedreicosaedre
   Call ActuGrille
End Sub



Private Sub mnuQuitter_Click()
   FenetrePolyedre.Hide
End Sub

Private Sub octaedre()
   NbSom% = 6     ' Nombre de sommets
   NbFacMax% = 8     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 1
   Ysom(1) = 0
   Zsom(1) = 0
   Xsom(2) = 0
   Ysom(2) = 1
   Zsom(2) = 0
   Xsom(3) = -1
   Ysom(3) = 0
   Zsom(3) = 0
   Xsom(4) = 0
   Ysom(4) = -1
   Zsom(4) = 0
   Xsom(5) = 0
   Ysom(5) = 0
   Zsom(5) = 1
   Xsom(6) = 0
   Ysom(6) = 0
   Zsom(6) = -1
   ' *****************************************************
   ' ***************  valeurs des arêtes : ***************
   ' *****************************************************
   Arete%(1, 2) = 1
   Arete%(1, 3) = 0
   Arete%(1, 4) = 1
   Arete%(1, 5) = 1
   Arete%(1, 6) = 2
   Arete%(2, 3) = 1
   Arete%(2, 4) = 0
   Arete%(2, 5) = 1
   Arete%(2, 6) = 2
   Arete%(3, 4) = 1
   Arete%(3, 5) = 1
   Arete%(3, 6) = 2
   Arete%(4, 5) = 1
   Arete%(4, 6) = 2
   ' *****************************************************
   RayonMax = 1
   ' *****************************************************
End Sub




Private Sub mnuRhombododecaedre_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre particulier"
   NomPolyedre$ = "Rhombododécaèdre"
   lblNomPol.Caption = NomPolyedre$
   Call rhombododecaedre
   Call ActuGrille
End Sub

Private Sub mnuTetraedre_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre régulier"
   NomPolyedre$ = "Tétraèdre"
   lblNomPol.Caption = NomPolyedre$
   Call tetraedre
   Call ActuGrille
End Sub


Private Sub mnuTetragonoTriOctaedre_Click()
   DefiniUt = False
   mnuOuvrir.Enabled = False
   lblTypePol.Caption = "Polyèdre particulier"
   NomPolyedre$ = "Tétragono-tri-octaèdre"
   lblNomPol.Caption = NomPolyedre$
   Call tetragonotrioctaedre
   Call ActuGrille
End Sub

Private Sub mnuVisu_Click()
   Call Visu
End Sub


Private Sub pentagonododecaedre()
   NbSom% = 20     ' Nombre de sommets
   NbFacMax% = 36     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 0
   Ysom(1) = UAT
   Zsom(1) = TAU
   Xsom(2) = 0
   Ysom(2) = -UAT
   Zsom(2) = TAU
   Xsom(3) = 1
   Ysom(3) = 1
   Zsom(3) = 1
   Xsom(4) = -1
   Ysom(4) = 1
   Zsom(4) = 1
   Xsom(5) = -1
   Ysom(5) = -1
   Zsom(5) = 1
   Xsom(6) = 1
   Ysom(6) = -1
   Zsom(6) = 1
   Xsom(7) = TAU
   Ysom(7) = 0
   Zsom(7) = UAT
   Xsom(8) = -TAU
   Ysom(8) = 0
   Zsom(8) = UAT
   Xsom(9) = UAT
   Ysom(9) = TAU
   Zsom(9) = 0
   Xsom(10) = -UAT
   Ysom(10) = TAU
   Zsom(10) = 0
   Xsom(11) = -UAT
   Ysom(11) = -TAU
   Zsom(11) = 0
   Xsom(12) = UAT
   Ysom(12) = -TAU
   Zsom(12) = 0
   Xsom(13) = TAU
   Ysom(13) = 0
   Zsom(13) = -UAT
   Xsom(14) = -TAU
   Ysom(14) = 0
   Zsom(14) = -UAT
   Xsom(15) = 1
   Ysom(15) = 1
   Zsom(15) = -1
   Xsom(16) = -1
   Ysom(16) = 1
   Zsom(16) = -1
   Xsom(17) = -1
   Ysom(17) = -1
   Zsom(17) = -1
   Xsom(18) = 1
   Ysom(18) = -1
   Zsom(18) = -1
   Xsom(19) = 0
   Ysom(19) = UAT
   Zsom(19) = -TAU
   Xsom(20) = 0
   Ysom(20) = -UAT
   Zsom(20) = -TAU
   ' ****************************************************
End Sub

Private Sub pentagonododecaedreicosaedre()
   NbSom% = 30     ' Nombre de sommets
   NbFacMax% = 60     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 0
   Ysom(1) = 0
   Zsom(1) = -2
   Xsom(2) = UAT
   Ysom(2) = 1
   Zsom(2) = -TAU
   Xsom(3) = -UAT
   Ysom(3) = 1
   Zsom(3) = -TAU
   Xsom(4) = -UAT
   Ysom(4) = -1
   Zsom(4) = -TAU
   Xsom(5) = UAT
   Ysom(5) = -1
   Zsom(5) = -TAU
   Xsom(6) = TAU
   Ysom(6) = -UAT
   Zsom(6) = -1
   Xsom(7) = TAU
   Ysom(7) = UAT
   Zsom(7) = -1
   Xsom(8) = 1
   Ysom(8) = TAU
   Zsom(8) = -UAT
   Xsom(9) = -1
   Ysom(9) = TAU
   Zsom(9) = -UAT
   Xsom(10) = -TAU
   Ysom(10) = UAT
   Zsom(10) = -1
   Xsom(11) = -TAU
   Ysom(11) = -UAT
   Zsom(11) = -1
   Xsom(12) = -1
   Ysom(12) = -TAU
   Zsom(12) = -UAT
   Xsom(13) = 1
   Ysom(13) = -TAU
   Zsom(13) = -UAT
   Xsom(14) = 2
   Ysom(14) = 0
   Zsom(14) = 0
   Xsom(15) = 0
   Ysom(15) = 2
   Zsom(15) = 0
   Xsom(16) = -2
   Ysom(16) = 0
   Zsom(16) = 0
   Xsom(17) = 0
   Ysom(17) = -2
   Zsom(17) = 0
   Xsom(18) = 1
   Ysom(18) = -TAU
   Zsom(18) = UAT
   Xsom(19) = TAU
   Ysom(19) = -UAT
   Zsom(19) = 1
   Xsom(20) = TAU
   Ysom(20) = UAT
   Zsom(20) = 1
   Xsom(21) = 1
   Ysom(21) = TAU
   Zsom(21) = UAT
   Xsom(22) = -1
   Ysom(22) = TAU
   Zsom(22) = UAT
   Xsom(23) = -TAU
   Ysom(23) = UAT
   Zsom(23) = 1
   Xsom(24) = -TAU
   Ysom(24) = -UAT
   Zsom(24) = 1
   Xsom(25) = -1
   Ysom(25) = -TAU
   Zsom(25) = UAT
   Xsom(26) = UAT
   Ysom(26) = -1
   Zsom(26) = TAU
   Xsom(27) = UAT
   Ysom(27) = 1
   Zsom(27) = TAU
   Xsom(28) = -UAT
   Ysom(28) = 1
   Zsom(28) = TAU
   Xsom(29) = -UAT
   Ysom(29) = -1
   Zsom(29) = TAU
   Xsom(30) = 0
   Ysom(30) = 0
   Zsom(30) = 2
   ' ****************************************************
End Sub

Private Sub PolReDim()
   ' -------------- ReDims pour polyèdres ----------------
   ' Coordonnées de sommets :
   ReDim Xsom(1 To NbSom%), Ysom(1 To NbSom%), Zsom(1 To NbSom%)
   ' Arêtes :
   ' 0 : pas d'arête; 1 : arête visible; 2 : arête invisible
   ReDim Arete%(1 To NbSom%, 1 To NbSom%)
   ' Rayon
   ReDim Rayon(1 To NbSom%)
   ' Numéro du point constituant
   ' le NSième sommet de la face JF
   ReDim NumSomFac%(1 To NbFacMax%, 1 To NbSom%)
   ' signe +1 ou -1 à appliquer au produit vectoriel V des
   ' 2 premières arêtes de la face JF pour qu'il soit dirigé
   ' vers l'extérieur du polyèdre
   ReDim SignePV%(1 To NbFacMax%)
   ' nombre de sommets de la face JF
   ReDim NbSomFac%(1 To NbFacMax%)
   ' coordonnées du centre de gravité de la face JF
   ReDim Xfac(1 To NbFacMax%), Yfac(1 To NbFacMax%), Zfac(1 To NbFacMax%)
   '------------------------------------------------------
End Sub

Private Sub rhombododecaedre()
   NbSom% = 14     ' Nombre de sommets
   NbFacMax% = 24     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 0
   Ysom(1) = 0
   Zsom(1) = 1
   Xsom(2) = 0.5
   Ysom(2) = 0.5
   Zsom(2) = 0.5
   Xsom(3) = -0.5
   Ysom(3) = 0.5
   Zsom(3) = 0.5
   Xsom(4) = -0.5
   Ysom(4) = -0.5
   Zsom(4) = 0.5
   Xsom(5) = 0.5
   Ysom(5) = -0.5
   Zsom(5) = 0.5
   Xsom(6) = 1
   Ysom(6) = 0
   Zsom(6) = 0
   Xsom(7) = 0
   Ysom(7) = 1
   Zsom(7) = 0
   Xsom(8) = -1
   Ysom(8) = 0
   Zsom(8) = 0
   Xsom(9) = 0
   Ysom(9) = -1
   Zsom(9) = 0
   Xsom(10) = 0.5
   Ysom(10) = 0.5
   Zsom(10) = -0.5
   Xsom(11) = -0.5
   Ysom(11) = 0.5
   Zsom(11) = -0.5
   Xsom(12) = -0.5
   Ysom(12) = -0.5
   Zsom(12) = -0.5
   Xsom(13) = 0.5
   Ysom(13) = -0.5
   Zsom(13) = -0.5
   Xsom(14) = 0
   Ysom(14) = 0
   Zsom(14) = -1
   ' ****************************************************
End Sub

Private Sub tetraedre()
   NbSom% = 4     ' Nombre de sommets
   NbFacMax% = 4     ' Nombre de faces
   Call PolReDim
   NbAr% = 6
   NbFac% = 4
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 1
   Ysom(1) = 1
   Zsom(1) = 1
   Xsom(2) = -1
   Ysom(2) = -1
   Zsom(2) = 1
   Xsom(3) = 1
   Ysom(3) = -1
   Zsom(3) = -1
   Xsom(4) = -1
   Ysom(4) = 1
   Zsom(4) = -1
   ' *****************************************************
   ' ***************  valeurs des arêtes : ***************
   ' *****************************************************
   Arete%(1, 2) = 1
   Arete%(1, 3) = 1
   Arete%(1, 4) = 1
   Arete%(2, 3) = 1
   Arete%(2, 4) = 1
   Arete%(3, 4) = 2
   ' *****************************************************
   NbSomFac%(1) = 3
   NbSomFac%(2) = 3
   NbSomFac%(3) = 3
   ' *****************************************************
   NumSomFac%(1, 1) = 1
   NumSomFac%(1, 2) = 2
   NumSomFac%(1, 3) = 3
   NumSomFac%(2, 1) = 1
   NumSomFac%(2, 2) = 2
   NumSomFac%(2, 3) = 4
   NumSomFac%(3, 1) = 1
   NumSomFac%(3, 2) = 3
   NumSomFac%(3, 3) = 4
   NumSomFac%(4, 1) = 2
   NumSomFac%(4, 2) = 3
   NumSomFac%(4, 3) = 4
   ' *****************************************************
   SignePV%(1) = 1
   SignePV%(2) = -1
   SignePV%(3) = 1
   SignePV%(4) = -1
   ' *****************************************************
   RayonMax = Sqr(3)
   ' *****************************************************
End Sub

Private Sub tetragonotrioctaedre()
   NbSom% = 26     ' Nombre de sommets
   NbFacMax% = 48     ' Nombre de faces
   Call PolReDim
   ' ****************************************************
   ' ******  valeurs des coordonnées des sommets : ******
   ' ****************************************************
   Xsom(1) = 0
   Ysom(1) = 0
   Zsom(1) = 1
   Xsom(2) = 0.666666667
   Ysom(2) = 0
   Zsom(2) = 0.666666667
   Xsom(3) = 0.5
   Ysom(3) = 0.5
   Zsom(3) = 0.5
   Xsom(4) = 0
   Ysom(4) = 0.666666667
   Zsom(4) = 0.666666667
   Xsom(5) = -0.5
   Ysom(5) = 0.5
   Zsom(5) = 0.5
   Xsom(6) = -0.666666667
   Ysom(6) = 0
   Zsom(6) = 0.666666667
   Xsom(7) = -0.5
   Ysom(7) = -0.5
   Zsom(7) = 0.5
   Xsom(8) = 0
   Ysom(8) = -0.666666667
   Zsom(8) = 0.666666667
   Xsom(9) = 0.5
   Ysom(9) = -0.5
   Zsom(9) = 0.5
   Xsom(10) = 1
   Ysom(10) = 0
   Zsom(10) = 0
   Xsom(11) = 0.666666667
   Ysom(11) = 0.666666667
   Zsom(11) = 0
   Xsom(12) = 0
   Ysom(12) = 1
   Zsom(12) = 0
   Xsom(13) = -0.666666667
   Ysom(13) = 0.666666667
   Zsom(13) = 0
   Xsom(14) = -1
   Ysom(14) = 0
   Zsom(14) = 0
   Xsom(15) = -0.666666667
   Ysom(15) = -0.666666667
   Zsom(15) = 0
   Xsom(16) = 0
   Ysom(16) = -1
   Zsom(16) = 0
   Xsom(17) = 0.666666667
   Ysom(17) = -0.666666667
   Zsom(17) = 0
   Xsom(18) = 0.666666667
   Ysom(18) = 0
   Zsom(18) = -0.666666667
   Xsom(19) = 0.5
   Ysom(19) = 0.5
   Zsom(19) = -0.5
   Xsom(20) = 0
   Ysom(20) = 0.666666667
   Zsom(20) = -0.666666667
   Xsom(21) = -0.5
   Ysom(21) = 0.5
   Zsom(21) = -0.5
   Xsom(22) = -0.666666667
   Ysom(22) = 0
   Zsom(22) = -0.666666667
   Xsom(23) = -0.5
   Ysom(23) = -0.5
   Zsom(23) = -0.5
   Xsom(24) = 0
   Ysom(24) = -0.666666667
   Zsom(24) = -0.666666667
   Xsom(25) = 0.5
   Ysom(25) = -0.5
   Zsom(25) = -0.5
   Xsom(26) = 0
   Ysom(26) = 0
   Zsom(26) = -1
   ' ****************************************************
End Sub

Private Sub txtNbSom_Change()
   If DefiniUt = True Then ' si polyèdre quelconque
      NbSom% = CInt(txtNbSom.Text)
      GridSom.Rows = NbSom% + 1
      ' ********************************
      ' renumérotation 1ère colonne
      GridSom.Col = 0
      For i% = 1 To NbSom%
         GridSom.Row = i%
         GridSom.Text = Format(i%, "0")
      Next i%
      ' ********************************
   Else
      txtNbSom.Text = Format(NbSom%, "0")
   End If
End Sub


Public Sub Visu()
   On Error GoTo Traite_ErreursVisuPol
   If DefiniUt = True Then ' si polyèdre quelconque
      If NbSom% < 4 Then
         Beep
         MsgBox "Un polyèdre comporte au moins 4 sommets !", 48, "POLYEDRE"
         Exit Sub
      End If
      NbFacMax% = 2 * NbSom% - 4
      Call PolReDim
      ' ********************************************
      ' affectation de leurs valeurs aux coordonnées
      ' ********************************************
      For i% = 1 To NbSom%
         FenetrePolyedre.GridSom.Row = i%
         FenetrePolyedre.GridSom.Col = 1
         Xsom(i%) = CSng(FenetrePolyedre.GridSom.Text)
         FenetrePolyedre.GridSom.Col = 2
         Ysom(i%) = CSng(FenetrePolyedre.GridSom.Text)
         FenetrePolyedre.GridSom.Col = 3
         Zsom(i%) = CSng(FenetrePolyedre.GridSom.Text)
      Next i%
      ' ********************************************
   End If
   FenetreRotaPol.Show
   Exit Sub
Traite_ErreursVisuPol:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Public Sub ActuGrille()
   txtNbSom.Text = CStr(NbSom%)
   GridSom.Rows = NbSom% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   GridSom.Col = 0
   For i% = 1 To NbSom%
      GridSom.Row = i%
      GridSom.Text = Format(i%, "0")
   Next i%
   ' ***************************************
   ' ***** actualisation de la grille *****
   ' ***************************************
   For i% = 1 To NbSom%
      FenetrePolyedre.GridSom.Row = i%
      FenetrePolyedre.GridSom.Col = 1
      FenetrePolyedre.GridSom.Text = Format(Xsom(i%), "0.000000000")
      FenetrePolyedre.GridSom.Col = 2
      FenetrePolyedre.GridSom.Text = Format(Ysom(i%), "0.000000000")
      FenetrePolyedre.GridSom.Col = 3
      FenetrePolyedre.GridSom.Text = Format(Zsom(i%), "0.000000000")
   Next i%
   ' **************************************
End Sub
