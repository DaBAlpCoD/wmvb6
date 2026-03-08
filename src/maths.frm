VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Maths 
   Caption         =   "Maths"
   ClientHeight    =   7935
   ClientLeft      =   195
   ClientTop       =   735
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "maths.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7935
   ScaleMode       =   0  'User
   ScaleWidth      =   10669.93
   Begin VB.ListBox ListMath 
      BackColor       =   &H80000000&
      Height          =   7665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
   Begin MSComDlg.CommonDialog ctrlCMDialog 
      Left            =   240
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuOK 
      Caption         =   "&OK"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
   Begin VB.Menu mnuAide 
      Caption         =   "&?"
      Begin VB.Menu mnuAPropos 
         Caption         =   "&A propos..."
      End
   End
End
Attribute VB_Name = "Maths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
   '--------------------------------------------------------
   '--------------------- Liste du menu principal -------------
   '--------------------------------------------------------
   ListMath.FontUnderline = False
   ListMath.ForeColor = NOIR
   ListMath.List(NUMCALC) = "Calculatrice"
   ListMath.List(NUMBLANC1) = ""
   ListMath.ForeColor = ROUGE
   ListMath.FontUnderline = True
   ListMath.List(NUMARITHM) = "ARITHMETIQUE :"
   ListMath.FontUnderline = False
   ListMath.ForeColor = NOIR
   ListMath.List(NUMBLANC2) = ""
   ListMath.List(NUMDECFAPRE) = "Décomposition d'un nombre entier en produit de facteurs premiers"
   ListMath.List(NUMPPCMPGCD) = "Plus Petit Commun Multiple et Plus Grand Commun Diviseur de deux nombres entiers"
   ListMath.List(NUMNOMBPREM) = "Liste des nombres premiers"
   ListMath.List(NUMTRINOMB) = "Tri de nombres"
   ListMath.List(NUMCOMPLEXES) = "Nombres complexes"
   ListMath.List(NUMBLANC3) = ""
   ListMath.ForeColor = ROUGE
   ListMath.Font.underline = True
   ListMath.List(NUMALGEBRE) = "ALGEBRE :"
   ListMath.Font.underline = False
   ListMath.ForeColor = NOIR
   ListMath.List(NUMBLANC4) = ""
   ListMath.List(NUMPOLYNOME) = "Polynômes"
   ListMath.List(NUMSOLEQ) = "Solution de l'équation F(X)=0"
   ListMath.ForeColor = BLEU
   ListMath.List(NUMALGLIN) = "ALGEBRE LINEAIRE :"
   ListMath.ForeColor = NOIR
   ListMath.List(NUMATRICE) = "Calculs sur les matrices"
   ListMath.List(NUMSYSLIN) = "Système linéaire de n équations à n inconnues"
   ListMath.List(NUMSYSNONLIN) = "Système non linéaire de n équations à n inconnues"
   ListMath.List(NUMPROLIN) = "Programmation linéaire (obtimisation d'une fonction objectif sous certaines contraintes)"
   ListMath.List(NUMBLANC5) = ""
   ListMath.ForeColor = ROUGE
   ListMath.Font.underline = True
   ListMath.List(NUMANALYSE) = "ANALYSE :"
   ListMath.Font.underline = False
   ListMath.ForeColor = NOIR
   ListMath.List(NUMBLANC6) = ""
   ListMath.List(NUMCOUR) = "Tracé d'une courbe"
   ListMath.List(NUMDERIVE) = "Dérivée d'une expression numérique"
   ListMath.List(NUMEQUADIF) = "Résolution graphique d'une équation différentielle du premier ordre du type dY/dX = F(X,Y)"
   ListMath.List(NUMEQUADIF2) = "Résolution graphique d'une équation différentielle du second ordre du type d²Y/dX² = F(X,Y,dY/dX)"
   ListMath.List(NUMPRIMITIVE) = "Tracé de la primitive d'une fonction"
   ListMath.List(NUMINTEGRALE) = "Intégrale d'une fonction"
   ListMath.ForeColor = BLEU
   ListMath.List(NUMINTERAP) = "INTERPOLATION ET APPROXIMATION :"
   ListMath.ForeColor = NOIR
   ListMath.List(NUMOINCAR) = "Ajustement d'une courbe à un nuage de points (affinement par moindres carrés)"
   ListMath.List(NUMREGRELIN) = "Regression linéaire (ajustement d'une droite à un nuage de points)"
   ListMath.List(NUMINPOLAG) = "Passage d'une courbe par un nuage de points (polynôme de Lagrange)"
   ListMath.List(NUMSPLINE) = "Passage d'une courbe par un nuage de points (Spline)"
   ListMath.List(NUMBLANC7) = ""
   ListMath.ForeColor = ROUGE
   ListMath.Font.underline = True
   ListMath.List(NUMGEOM) = "GEOMETRIE :"
   ListMath.Font.underline = False
   ListMath.ForeColor = NOIR
   ListMath.List(NUMBLANC8) = ""
   ListMath.List(NUMTRIANGLE) = "Triangles"
   ListMath.List(NUMPOLYEDRE) = "Polyèdres"
   ListMath.List(NUMSURF3D) = "Surface 3D"
   ListMath.Selected(NUMCOUR) = True
   '---------------------------------------------------------------
   '--------------------- Fonctions numériques --------------------
   '---------------------------------------------------------------
   FoNum$(1) = "ABS("
   FoNum$(2) = "ATN("
   FoNum$(3) = "COS("
   FoNum$(4) = "EXP("
   FoNum$(5) = "FIX("
   FoNum$(6) = "INT("
   FoNum$(7) = "LOG("
   FoNum$(8) = "SGN("
   FoNum$(9) = "SIN("
   FoNum$(10) = "SQR("
   FoNum$(11) = "TAN("
   '---------------------------------------------------------------
   '--------------------- Initialisations -------------------------
   '---------------------------------------------------------------
   NbPts% = 0
   '---------------------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub




Private Sub ListMath_DblClick()
   mnuOK_Click
End Sub


Private Sub mnuAPropos_Click()
   APropos.Show
End Sub

Private Sub mnuOK_Click()
NumChoisi% = ListMath.ListIndex
   Select Case NumChoisi%
   Case NUMCALC
      lancer = Shell("calc.exe", 1)
   Case NUMDECFAPRE
      FenetreDecFaPre.Show
   Case NUMPPCMPGCD
      FenetrePPCMPGCD.Show
   Case NUMNOMBPREM
      FenetreNbPremiers.Show
   Case NUMTRINOMB
      FenetreTriNombres.Show
   Case NUMCOMPLEXES
      FenetreComplexe.Show
   Case NUMCOUR
      FenetreDefCourbe.Show
   Case NUMSOLEQ
      FenetreSolEq.Show
   Case NUMEQUADIF
      FenetreDefEqDif.Show
   Case NUMEQUADIF2
      FenetreDefEqDif2.Show
   Case NUMPRIMITIVE
      FenetreDefPrim.Show
   Case NUMINTEGRALE
      FenetreIntegrale.Show
   Case NUMDERIVE
      FenetreDerive.Show
   Case NUMOINCAR
      FenetreDefMoinCar.Show
   Case NUMREGRELIN
      FenetreDefRegreLin.Show
   Case NUMINPOLAG
      FenetreDefInPoLag.Show
   Case NUMSPLINE
      FenetreDefSpline.Show
   Case NUMPOLYNOME
      FenetrePolynome.Show
   Case NUMSYSLIN
      FenetreSysLin.Show
   Case NUMSYSNONLIN
      FenetreSysNonLin.Show
   Case NUMATRICE
      FenetreMatrice.Show
   Case NUMPROLIN
      FenetreProLin.Show
   Case NUMTRIANGLE
      FenetreTriangle.Show
   Case NUMPOLYEDRE
      FenetrePolyedre.Show
   Case NUMSURF3D
      FenetreSurf3D.Show
   End Select
End Sub



Private Sub mnuQuitter_Click()
   Unload Maths
End Sub


