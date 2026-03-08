VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreMatrice 
   Caption         =   "Matrices carrées"
   ClientHeight    =   6615
   ClientLeft      =   1485
   ClientTop       =   1455
   ClientWidth     =   9990
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
   ScaleWidth      =   9990
   Begin MSFlexGridLib.MSFlexGrid GridPImat 
      Height          =   1095
      Left            =   5280
      TabIndex        =   16
      Top             =   4440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
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
   Begin MSFlexGridLib.MSFlexGrid GridPRmat 
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1931
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
   Begin MSFlexGridLib.MSFlexGrid GridPmat 
      Height          =   1215
      Left            =   5280
      TabIndex        =   14
      Top             =   2760
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSFlexGridLib.MSFlexGrid GridWmat 
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
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
   Begin MSFlexGridLib.MSFlexGrid GridM2mat 
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSFlexGridLib.MSFlexGrid GridMmat 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
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
   Begin VB.TextBox txtOrdreMat 
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
      Text            =   "3"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblPImat 
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label lblPRmat 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label lblValDet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   9495
   End
   Begin VB.Label lblPmat 
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblM2mat 
      Caption         =   "Matrice M2 :"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDetMmat 
      Caption         =   "Déterminant de M :"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblWmat 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblOrdreMat 
      Caption         =   "Ordre des matrices :"
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
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblMmat 
      Caption         =   "Matrice M :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
         Begin VB.Menu mnuOuvrirM 
            Caption         =   "Matrice &M"
         End
         Begin VB.Menu mnuOuvrirM2 
            Caption         =   "Matrice M2"
         End
      End
      Begin VB.Menu mnuEnregistrer 
         Caption         =   "&Enregistrer..."
         Begin VB.Menu mnuEnregM 
            Caption         =   "Matrice M"
         End
         Begin VB.Menu mnuEnregM2 
            Caption         =   "Matrice M2"
         End
         Begin VB.Menu mnuEnregW 
            Caption         =   "Matrice W inverse de M"
         End
         Begin VB.Menu mnuEnregP 
            Caption         =   "Matrice P produit"
         End
         Begin VB.Menu mnuEnregVap 
            Caption         =   "Valeurs propres de M"
         End
         Begin VB.Menu mnuEnregVep 
            Caption         =   "Matrice des vecteurs propres de M"
         End
      End
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer..."
      Begin VB.Menu mnuProduitMM2 
         Caption         =   "Produit MxM2"
      End
      Begin VB.Menu mnuProduitWM2 
         Caption         =   "Produit WxM2"
      End
      Begin VB.Menu mnuProduitWM2M 
         Caption         =   "Produit WxM2xM"
      End
      Begin VB.Menu mnuInversion 
         Caption         =   "&Inversion de M..."
         Begin VB.Menu mnuInvMet1 
            Caption         =   "Méthode n°&1 : méthode exacte passant par le calcul du déterminant. Limitée aux ordres inférieurs à 8."
         End
         Begin VB.Menu mnuInvMet2 
            Caption         =   "Méthode n°&2 : Méthode modifiée d'élimination de GAUSS-JORDAN."
         End
         Begin VB.Menu mnuInvMet3 
            Caption         =   "Méthode n°&3 : Méthode itérative passant par le calcul du polynôme caractéristique et du déterminant."
         End
         Begin VB.Menu mnuInvMet4 
            Caption         =   "Méthode n°&4 : Méthode de CHOLEVSKI (pour une matrice symétrique définie positive)."
         End
      End
      Begin VB.Menu mnuDiagonalisation 
         Caption         =   "&Diagonalisation de M..."
         Begin VB.Menu mnuDiagMet1 
            Caption         =   "Méthode n°1 : Méthode du double QR avec déplacement."
         End
         Begin VB.Menu mnuDiagMet2 
            Caption         =   "Méthode n°2 : Détermination du polynôme caractéristique (méthode de SOURIAU), recherche de ses zéros"
         End
         Begin VB.Menu mnuDiagMet2bis 
            Caption         =   " (méthode de BAIRSTOW) et détermination des vecteurs propres par résolution de systèmes sur-déterminés."
         End
         Begin VB.Menu mnuDiagMet3 
            Caption         =   "Méthode n°3 : Méthode de JACOBI classique (pour une matrice symétrique)."
         End
      End
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreMatrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumMethodInv%
Dim NumMethodDiag%
Dim NomMat$
Private Sub Form_Load()
   '***************************************
   OrdreMat% = 3
   '***************************************
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Wmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim M1mat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim M2mat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Pmat(1 To OrdreMat%, 1 To OrdreMat%)
   '***************************************
   '***** mise en place des 4 grilles *****
   '***************************************
   GridMmat.Rows = OrdreMat% + 1
   GridMmat.Cols = OrdreMat% + 1
   GridWmat.Rows = OrdreMat% + 1
   GridWmat.Cols = OrdreMat% + 1
   GridM2mat.Rows = OrdreMat% + 1
   GridM2mat.Cols = OrdreMat% + 1
   GridPmat.Rows = OrdreMat% + 1
   GridPmat.Cols = OrdreMat% + 1
   GridPRmat.Rows = OrdreMat% + 1
   GridPRmat.Cols = OrdreMat% + 1
   GridPImat.Rows = OrdreMat% + 1
   GridPImat.Cols = OrdreMat% + 1
   '***************************************
   For i% = 0 To OrdreMat%
      FenetreMatrice.GridMmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridMmat.ColWidth(i%) = 1000
      FenetreMatrice.GridWmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridWmat.ColWidth(i%) = 1000
      FenetreMatrice.GridM2mat.FixedAlignment(i%) = 2
      FenetreMatrice.GridM2mat.ColWidth(i%) = 1000
      FenetreMatrice.GridPmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPmat.ColWidth(i%) = 1000
      FenetreMatrice.GridPRmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPRmat.ColWidth(i%) = 1000
      FenetreMatrice.GridPImat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPImat.ColWidth(i%) = 1000
   Next i%
   ' ********************************
   ' numérotation 1ères lignes
   ' ********************************
   GridMmat.Row = 0
   GridWmat.Row = 0
   GridM2mat.Row = 0
   GridPmat.Row = 0
   GridPRmat.Row = 0
   GridPImat.Row = 0
   For i% = 1 To OrdreMat%
      GridMmat.Col = i%
      GridMmat.Text = i%
      GridWmat.Col = i%
      GridWmat.Text = i%
      GridM2mat.Col = i%
      GridM2mat.Text = i%
      GridPmat.Col = i%
      GridPmat.Text = i%
      GridPRmat.Col = i%
      GridPRmat.Text = i%
      GridPImat.Col = i%
      GridPImat.Text = i%
   Next i%
   ' *********************************
   ' numérotation 1ères colonnes
   ' *********************************
   GridMmat.Col = 0
   GridWmat.Col = 0
   GridM2mat.Col = 0
   GridPmat.Col = 0
   GridPRmat.Col = 0
   GridPImat.Col = 0
   For i% = 1 To OrdreMat%
      GridMmat.Row = i%
      GridMmat.Text = i%
      GridWmat.Row = i%
      GridWmat.Text = i%
      GridM2mat.Row = i%
      GridM2mat.Text = i%
      GridPmat.Row = i%
      GridPmat.Text = i%
      GridPRmat.Row = i%
      GridPRmat.Text = i%
      GridPImat.Row = i%
      GridPImat.Text = i%
   Next i%
   ' ****************************************
   '       Valeurs par défaut
   ' ****************************************
   Mmat(1, 1) = -4
   Mmat(1, 2) = 1
   Mmat(1, 3) = 3
   Mmat(2, 1) = 1
   Mmat(2, 2) = 2
   Mmat(2, 3) = 2
   Mmat(3, 1) = 3
   Mmat(3, 2) = 2
   Mmat(3, 3) = 1
   '***************************************
   M2mat(1, 1) = 1
   M2mat(1, 2) = 1
   M2mat(1, 3) = 3
   M2mat(2, 1) = -1
   M2mat(2, 2) = 2
   M2mat(2, 3) = 1
   M2mat(3, 1) = 2
   M2mat(3, 2) = 2
   M2mat(3, 3) = 1
   ' ****************************************
   ' placement des valeurs par défaut de Mmat
   ' ****************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridMmat.Row = i%
         FenetreMatrice.GridMmat.Col = j%
         FenetreMatrice.GridMmat.Text = Format(Mmat(i%, j%), "0.000")
      Next j%
   Next i%
   ' ****************************************
   ' placement des valeurs par défaut de M2mat
   ' ****************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridM2mat.Row = i%
         FenetreMatrice.GridM2mat.Col = j%
         FenetreMatrice.GridM2mat.Text = Format(M2mat(i%, j%), "0.000")
      Next j%
   Next i%
   ' ****************************************
End Sub

Private Sub gridM2mat_KeyPress(KeyAscii As Integer)
   EleMText$ = GridM2mat.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      GridM2mat.Text = EleMText$
   Case 8
      If Len(GridM2mat.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         GridM2mat.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub gridMmat_KeyPress(KeyAscii As Integer)
   EleMText$ = GridMmat.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      GridMmat.Text = EleMText$
   Case 8
      If Len(GridMmat.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         GridMmat.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub InverseMat()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Mmat, Wmat
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Wmat(1 To OrdreMat%, 1 To OrdreMat%)
   ' *************************************************
   '               Initialisations
   ' *************************************************
   lblValDet.Caption = ""
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' *************************************************
   ' remise en place de la grille de Wmat
   ' **************************************
   GridWmat.Rows = OrdreMat% + 1
   GridWmat.Cols = OrdreMat% + 1
   '***************************************
   For i% = 0 To OrdreMat%
      FenetreMatrice.GridWmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridWmat.ColWidth(i%) = 1000
   Next i%
   ' ********************************
   ' numérotation 1ère ligne
   ' ********************************
   GridWmat.Row = 0
   For i% = 1 To OrdreMat%
      GridWmat.Col = i%
      GridWmat.Text = i%
   Next i%
   ' *********************************
   ' numérotation 1ère colonne
   ' *********************************
   GridWmat.Col = 0
   For i% = 1 To OrdreMat%
      GridWmat.Row = i%
      GridWmat.Text = i%
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Mmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridMmat.Row = i%
         FenetreMatrice.GridMmat.Col = j%
         If FenetreMatrice.GridMmat.Text = "" Then
            FenetreMatrice.GridMmat.Text = 0
            Mmat(i%, j%) = 0
         Else
            Mmat(i%, j%) = CSng(FenetreMatrice.GridMmat.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   ' calcul de W matrice inverse de M
   ' *************************************************
   Select Case NumMethodInv%
   Case 1
      Call InvMat01
      lblValDet.Caption = Format(DetMmat, "0.000")  ' affichage du déterminant de Mmat
   Case 2
      Call InvMat02
   Case 3
      Call InvMat03
      lblValDet.Caption = Format(DetMmat, "0.000")  ' affichage du déterminant de Mmat
   Case 4
      Call VerifieMatSym(OrdreMat%, Mmat())
      If Erreur = True Then
         lblInfo.Caption = ""
         Exit Sub
      End If
      DePo = True
      Call InvMatCholeski(OrdreMat%, DePo, DetMmat, Mmat(), Wmat())
      If DePo = False Then
         Message$ = "La matrice n'est pas définie positive !"
         MsgBox Message$, 48, "InverseMat"
         lblInfo.Caption = ""
         Exit Sub
      End If
      lblValDet.Caption = Format(DetMmat, "0.000")  ' affichage du déterminant de Mmat
   End Select
   ' *************************************************
   ' affichage des éléments de Wmat
   ' *************************************************
   lblWmat.Caption = "Matrice W inverse de M :"
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridWmat.Row = i%
         FenetreMatrice.GridWmat.Col = j%
         FenetreMatrice.GridWmat.Text = Format(Wmat(i%, j%), "0.000")
      Next j%
   Next i%
   ' *************************************************
   ' effacement du nom de Pmat
   ' *************************************************
   lblPmat.Caption = ""
   '***************************************
   '************ Explications *************
   '***************************************
   Info$ = "Dernier calcul effectué : Inversion de M : "
   Select Case NumMethodInv%
   Case 1
      Info$ = Info$ & "Méthode n°1 : Méthode exacte passant par le calcul du déterminant. Limitée aux ordres inférieurs à 8."
   Case 2
      Info$ = Info$ & "Méthode n°2 : Méthode modifiée d'élimination de GAUSS-JORDAN."
   Case 3
      Info$ = Info$ & "Méthode n°3 : Méthode itérative passant par le calcul du polynôme caractéristique et du déterminant."
   Case 4
      Info$ = Info$ & "Méthode n°4 : Méthode de CHOLEVSKI (pour une matrice symétrique définie positive)."
   End Select
   lblInfo.ForeColor = BLEU
   lblInfo.Font.bold = True
   lblInfo.Font.underline = False
   lblInfo.Caption = Info$
   ' *************************************************
End Sub


Private Sub mnuDiagMet1_Click()
   NumMethodDiag% = 1
   Call DiaMat
End Sub


Private Sub mnuDiagMet2_Click()
   NumMethodDiag% = 2
   Call DiaMat
End Sub



Private Sub mnuDiagMet2bis_Click()
   NumMethodDiag% = 2
   Call DiaMat
End Sub


Private Sub mnuDiagMet3_Click()
   NumMethodDiag% = 3
   Call DiaMat
End Sub



Private Sub mnuEnregM_Click()
   NomMat$ = "M"
   Call EnregMat
End Sub


Private Sub mnuEnregM2_Click()
   NomMat$ = "M2"
   Call EnregMat
End Sub



Private Sub mnuEnregP_Click()
   NomMat$ = "P"
   Call EnregMat
End Sub






Private Sub mnuEnregVap_Click()
   NomMat$ = "Vap"
   Call EnregMat
End Sub

Private Sub mnuEnregVep_Click()
   NomMat$ = "Vep"
   Call EnregMat
End Sub


Private Sub mnuEnregW_Click()
   NomMat$ = "W"
   Call EnregMat
End Sub

Private Sub mnuInvMet1_Click()
   NumMethodInv% = 1
   Call InverseMat
End Sub


Private Sub mnuInvMet2_Click()
   NumMethodInv% = 2
   Call InverseMat
End Sub


Private Sub mnuInvMet3_Click()
   NumMethodInv% = 3
   Call InverseMat
End Sub







Private Sub mnuInvMet4_Click()
   NumMethodInv% = 4
   Call InverseMat
End Sub


Private Sub mnuOuvrirM_Click()
   NomMat$ = "M"
   Call OuvreMat
End Sub


Private Sub mnuOuvrirM2_Click()
   NomMat$ = "M2"
   Call OuvreMat
End Sub


Private Sub mnuProduitMM2_Click()
   Call ProduitMxM2
End Sub


Private Sub mnuProduitWM2_Click()
   Call ProduitWxM2
End Sub

Private Sub mnuProduitWM2M_Click()
    Call ProduitWxM2xM
End Sub

Private Sub mnuQuitter_Click()
   FenetreMatrice.Hide
End Sub

Private Sub ProduitMxM2()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Mmat, M2mat, Pmat
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim M2mat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Pmat(1 To OrdreMat%, 1 To OrdreMat%)
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Mmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridMmat.Row = i%
         FenetreMatrice.GridMmat.Col = j%
         If FenetreMatrice.GridMmat.Text = "" Then
            FenetreMatrice.GridMmat.Text = 0
            Mmat(i%, j%) = 0
         Else
            Mmat(i%, j%) = CSng(FenetreMatrice.GridMmat.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de M2mat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridM2mat.Row = i%
         FenetreMatrice.GridM2mat.Col = j%
         If FenetreMatrice.GridM2mat.Text = "" Then
            FenetreMatrice.GridM2mat.Text = 0
            M2mat(i%, j%) = 0
         Else
            M2mat(i%, j%) = CSng(FenetreMatrice.GridM2mat.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   ' redimentionnement de la grille gridPmat
   ' *************************************************
   GridPmat.Rows = OrdreMat% + 1
   GridPmat.Cols = OrdreMat% + 1
   '***************************************
   For i% = 0 To OrdreMat%
      FenetreMatrice.GridPmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPmat.ColWidth(i%) = 1000
   Next i%
   ' **********************************
   ' renumérotation 1ère ligne gridPmat
   ' **********************************
   GridPmat.Row = 0
   For i% = 1 To OrdreMat%
      GridPmat.Col = i%
      GridPmat.Text = i%
   Next i%
   ' ************************************
   ' renumérotation 1ère colonne gridPmat
   ' ************************************
   GridPmat.Col = 0
   For i% = 1 To OrdreMat%
      GridPmat.Row = i%
      GridPmat.Text = i%
   Next i%
   ' *************************************************
   ' calcul de P produit de M et M2
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         Pmat(i%, j%) = 0
         For k% = 1 To OrdreMat%
            Pmat(i%, j%) = Pmat(i%, j%) + Mmat(i%, k%) * M2mat(k%, j%)
         Next k%
      Next j%
   Next i%
   ' *************************************************
   'affichage du produit effectué
   ' *************************************************
   lblPmat.Caption = "Matrice P produit de M et M2 :"
   ' *************************************************
   ' affichage des éléments de Pmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridPmat.Row = i%
         FenetreMatrice.GridPmat.Col = j%
         FenetreMatrice.GridPmat.Text = Format(Pmat(i%, j%), "0.000")
      Next j%
   Next i%
   ' ********************************************
   lblInfo.Caption = ""
   ' ********************************************
   '***************************************
   '************ Explications *************
   '***************************************
   Info$ = "Dernier calcul effectué : Produit MxM2"
   lblInfo.ForeColor = BLEU
   lblInfo.Font.bold = True
   lblInfo.Font.underline = False
   lblInfo.Caption = Info$
   '***************************************
End Sub

Private Sub ProduitWxM2()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase M2mat, Pmat
   ReDim M2mat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Pmat(1 To OrdreMat%, 1 To OrdreMat%)
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   On Error Resume Next
   DoEvents
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de M2mat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridM2mat.Row = i%
         FenetreMatrice.GridM2mat.Col = j%
         If FenetreMatrice.GridM2mat.Text = "" Then
            FenetreMatrice.GridM2mat.Text = 0
            M2mat(i%, j%) = 0
         Else
            M2mat(i%, j%) = CSng(FenetreMatrice.GridM2mat.Text)
         End If
      Next j%
   Next i%
   On Error GoTo 0
   ' *************************************************
   ' calcul de P produit de W et M2
   ' *************************************************
   If GridWmat.Cols = 2 Then
      Message$ = "Produit impossible !"
      MsgBox Message$, 48, "ProduitWxM2"
      Exit Sub
   End If
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         Pmat(i%, j%) = 0
         For k% = 1 To OrdreMat%
            Pmat(i%, j%) = Pmat(i%, j%) + Wmat(i%, k%) * M2mat(k%, j%)
         Next k%
      Next j%
   Next i%
   ' *************************************************
   'affichage du produit effectué
   ' *************************************************
   lblPmat.Caption = "Matrice P produit de W et M2 :"
   ' *************************************************
   ' affichage des éléments de Pmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridPmat.Row = i%
         FenetreMatrice.GridPmat.Col = j%
         FenetreMatrice.GridPmat.Text = Format(Pmat(i%, j%), "0.000")
      Next j%
   Next i%
   ' ********************************************
   lblInfo.Caption = ""
   ' ********************************************
   '***************************************
   '************ Explications *************
   '***************************************
   Info$ = "Dernier calcul effectué : Produit WxM2"
   lblInfo.ForeColor = BLEU
   lblInfo.Font.bold = True
   lblInfo.Font.underline = False
   lblInfo.Caption = Info$
   '***************************************
End Sub




Private Sub txtOrdreMat_Change()
   If txtOrdreMat.Text = "" Then
      OrdreMat% = 1
   Else
      OrdreMat% = CInt(txtOrdreMat.Text)
   End If
   If OrdreMat% < 1 Then
      Beep
      MsgBox "l'ordre de M doit être supérieur ou égal à 1 !", 48, "MATRICE"
      OrdreMat% = 1
      txtOrdreMat.Text = "1"
   ElseIf OrdreMat% > 30000 Then
      Beep
      MsgBox "l'ordre de M doit être inférieur à 30000 !", 48, "MATRICE"
      OrdreMat% = 1
   End If
   GridMmat.Rows = OrdreMat% + 1
   GridMmat.Cols = OrdreMat% + 1
   GridWmat.Rows = OrdreMat% + 1
   GridWmat.Cols = OrdreMat% + 1
   GridM2mat.Rows = OrdreMat% + 1
   GridM2mat.Cols = OrdreMat% + 1
   GridPmat.Rows = OrdreMat% + 1
   GridPmat.Cols = OrdreMat% + 1
   GridPRmat.Rows = OrdreMat% + 1
   GridPRmat.Cols = OrdreMat% + 1
   GridPImat.Rows = OrdreMat% + 1
   GridPImat.Cols = OrdreMat% + 1
   '***************************************
   For i% = 0 To OrdreMat%
      FenetreMatrice.GridMmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridMmat.ColWidth(i%) = 1000
      FenetreMatrice.GridWmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridWmat.ColWidth(i%) = 1000
      FenetreMatrice.GridM2mat.FixedAlignment(i%) = 2
      FenetreMatrice.GridM2mat.ColWidth(i%) = 1000
      FenetreMatrice.GridPmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPmat.ColWidth(i%) = 1000
      FenetreMatrice.GridPRmat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPRmat.ColWidth(i%) = 1000
      FenetreMatrice.GridPImat.FixedAlignment(i%) = 2
      FenetreMatrice.GridPImat.ColWidth(i%) = 1000
   Next i%
   ' **********************************
   ' renumérotation 1ère ligne gridMmat
   ' **********************************
   GridMmat.Row = 0
   For i% = 1 To OrdreMat%
      GridMmat.Col = i%
      GridMmat.Text = ""
      GridMmat.Text = Format(i%, "0")
   Next i%
   ' ************************************
   ' renumérotation 1ère colonne gridMmat
   ' ************************************
   GridMmat.Col = 0
   For i% = 1 To OrdreMat%
      GridMmat.Row = i%
      GridMmat.Text = ""
      GridMmat.Text = Format(i%, "0")
   Next i%
   ' **********************************
   ' renumérotation 1ère ligne gridWmat
   ' **********************************
   GridWmat.Row = 0
   For i% = 1 To OrdreMat%
      GridWmat.Col = i%
      GridWmat.Text = ""
      GridWmat.Text = Format(i%, "0")
   Next i%
   ' ************************************
   ' renumérotation 1ère colonne gridWmat
   ' ************************************
   GridWmat.Col = 0
   For i% = 1 To OrdreMat%
      GridWmat.Row = i%
      GridWmat.Text = ""
      GridWmat.Text = Format(i%, "0")
   Next i%
   ' ***********************************
   ' renumérotation 1ère ligne gridM2mat
   ' ***********************************
   GridM2mat.Row = 0
   For i% = 1 To OrdreMat%
      GridM2mat.Col = i%
      GridM2mat.Text = ""
      GridM2mat.Text = Format(i%, "0")
   Next i%
   ' *************************************
   ' renumérotation 1ère colonne gridM2mat
   ' *************************************
   GridM2mat.Col = 0
   For i% = 1 To OrdreMat%
      GridM2mat.Row = i%
      GridM2mat.Text = ""
      GridM2mat.Text = Format(i%, "0")
   Next i%
   ' **********************************
   ' renumérotation 1ère ligne gridPmat
   ' **********************************
   GridPmat.Row = 0
   For i% = 1 To OrdreMat%
      GridPmat.Col = i%
      GridPmat.Text = ""
      GridPmat.Text = Format(i%, "0")
   Next i%
   ' ************************************
   ' renumérotation 1ère colonne gridPmat
   ' ************************************
   GridPmat.Col = 0
   For i% = 1 To OrdreMat%
      GridPmat.Row = i%
      GridPmat.Text = ""
      GridPmat.Text = Format(i%, "0")
   Next i%
   ' ***********************************
   ' renumérotation 1ère ligne gridPRmat
   ' ***********************************
   GridPRmat.Row = 0
   For i% = 1 To OrdreMat%
      GridPRmat.Col = i%
      GridPRmat.Text = ""
      GridPRmat.Text = Format(i%, "0")
   Next i%
   ' *************************************
   ' renumérotation 1ère colonne gridPRmat
   ' *************************************
   GridPRmat.Col = 0
   For i% = 1 To OrdreMat%
      GridPRmat.Row = i%
      GridPRmat.Text = ""
      GridPRmat.Text = Format(i%, "0")
   Next i%
   ' ***********************************
   ' renumérotation 1ère ligne gridPImat
   ' ***********************************
   GridPImat.Row = 0
   For i% = 1 To OrdreMat%
      GridPImat.Col = i%
      GridPImat.Text = ""
      GridPImat.Text = Format(i%, "0")
   Next i%
   ' *************************************
   ' renumérotation 1ère colonne gridPImat
   ' *************************************
   GridPImat.Col = 0
   For i% = 1 To OrdreMat%
      GridPImat.Row = i%
      GridPImat.Text = ""
      GridPImat.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub
Public Sub OuvreMat()
   '-----------------------------------------------------
   On Error GoTo Traite_ErreursOuvMat
   Maths.ctrlCMDialog.Filter = "Matrice (*.mat)|*.mat"
   ' nom de fichier et chemin doivent exister
   ' sinon apparait un message d'erreur spécifique
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   '-----------------------------------------------------
   ' Ouverture et lecture du fichier d'éléments de matrice
   ' et écriture de ces éléments dans Mmat(i%,j%)
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, OrdreMatLoc%
   ' ********************************************
   '               ReDims
   ' ********************************************
   If OrdreMatLoc% <> OrdreMat% Then
      OrdreMat% = OrdreMatLoc%
      Erase Mmat, M2mat
      ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
      ReDim M2mat(1 To OrdreMat%, 1 To OrdreMat%)
   End If
   ' *************************************************
   ' lecture des éléments de la matrice
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         If NomMat$ = "M" Then
            Input #1, Mmat(i%, j%)
         ElseIf NomMat$ = "M2" Then
            Input #1, M2mat(i%, j%)
         End If
      Next j%
   Next i%
   Close #1
   ' *************************************************
   ' placement de l'ordre des matrices, ce qui provoque
   ' le redimentionnement des grilles
   ' *************************************************
   txtOrdreMat.Text = Format(OrdreMat%, "0")
   ' *********************************************
   ' placement des nouvelles valeurs de la matrice
   ' *********************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         If NomMat$ = "M" Then
            GridMmat.Row = i%
            GridMmat.Col = j%
            GridMmat.Text = ""
            GridMmat.Text = Format(Mmat(i%, j%), "0.000")
         ElseIf NomMat$ = "M2" Then
            GridM2mat.Row = i%
            GridM2mat.Col = j%
            GridM2mat.Text = ""
            GridM2mat.Text = Format(M2mat(i%, j%), "0.000")
         End If
      Next j%
   Next i%
   ' ****************************************
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursOuvMat:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Public Sub EnregMat()
   On Error GoTo Traite_ErreursEnregMat
   Select Case NomMat$
   Case "M"
      ' *************************************************
      Maths.ctrlCMDialog.DefaultExt = "mat"
      Maths.ctrlCMDialog.Filter = "Matrice (*.mat)|*.mat"
      Maths.ctrlCMDialog.Flags = &H2&
      Maths.ctrlCMDialog.Action = 2
      ' ********************************************
      '               ReDim
      ' ********************************************
      Erase Mmat
      ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
      ' *******************************************************
      ' affectation de leurs valeurs aux éléments de la matrice
      ' *******************************************************
      For i% = 1 To OrdreMat%
         For j% = 1 To OrdreMat%
            FenetreMatrice.GridMmat.Row = i%
            FenetreMatrice.GridMmat.Col = j%
            If FenetreMatrice.GridMmat.Text = "" Then
               FenetreMatrice.GridMmat.Text = 0
               Mmat(i%, j%) = 0
            Else
               Mmat(i%, j%) = CSng(FenetreMatrice.GridMmat.Text)
            End If
         Next j%
      Next i%
      ' *************************************************
      '-----------------------------------------------------
      ' Création du fichier d'éléments de matrice
      ' et écriture de ces éléments dans le fichier
      '-----------------------------------------------------
      Open Maths.ctrlCMDialog.FileName For Output As #1
         Write #1, OrdreMat%
         For i% = 1 To OrdreMat%
            For j% = 1 To OrdreMat%
               Write #1, Mmat(i%, j%)
            Next j%
         Next i%
      Close #1
      ' *************************************************
   Case "M2"
      ' *************************************************
      Maths.ctrlCMDialog.DefaultExt = "mat"
      Maths.ctrlCMDialog.Filter = "Matrice (*.mat)|*.mat"
      Maths.ctrlCMDialog.Flags = &H2&
      Maths.ctrlCMDialog.Action = 2
      ' ********************************************
      '               ReDim
      ' ********************************************
      Erase M2mat
      ReDim M2mat(1 To OrdreMat%, 1 To OrdreMat%)
      ' *******************************************************
      ' affectation de leurs valeurs aux éléments de la matrice
      ' *******************************************************
      For i% = 1 To OrdreMat%
         For j% = 1 To OrdreMat%
            FenetreMatrice.GridM2mat.Row = i%
            FenetreMatrice.GridM2mat.Col = j%
            If FenetreMatrice.GridM2mat.Text = "" Then
               FenetreMatrice.GridM2mat.Text = 0
               M2mat(i%, j%) = 0
            Else
               M2mat(i%, j%) = CSng(FenetreMatrice.GridM2mat.Text)
            End If
         Next j%
      Next i%
      ' *************************************************
      '-----------------------------------------------------
      ' Création du fichier d'éléments de matrice
      ' et écriture de ces éléments dans le fichier
      '-----------------------------------------------------
      Open Maths.ctrlCMDialog.FileName For Output As #1
         Write #1, OrdreMat%
         For i% = 1 To OrdreMat%
            For j% = 1 To OrdreMat%
               Write #1, M2mat(i%, j%)
            Next j%
         Next i%
      Close #1
      ' *************************************************
   Case "P"
      If lblPmat.Caption = "Matrice P produit de M et M2 :" Or lblPmat.Caption = "Matrice P produit de W et M2 :" Then
         ' *************************************************
         Maths.ctrlCMDialog.DefaultExt = "mat"
         Maths.ctrlCMDialog.Filter = "Matrice (*.mat)|*.mat"
         Maths.ctrlCMDialog.Flags = &H2&
         Maths.ctrlCMDialog.Action = 2
         ' ********************************************
         '               ReDim
         ' ********************************************
         Erase Pmat
         ReDim Pmat(1 To OrdreMat%, 1 To OrdreMat%)
         ' *******************************************************
         ' affectation de leurs valeurs aux éléments de la matrice
         ' *******************************************************
         For i% = 1 To OrdreMat%
            For j% = 1 To OrdreMat%
               FenetreMatrice.GridPmat.Row = i%
               FenetreMatrice.GridPmat.Col = j%
               If FenetreMatrice.GridPmat.Text = "" Then
                  FenetreMatrice.GridPmat.Text = 0
                  Pmat(i%, j%) = 0
               Else
                  Pmat(i%, j%) = CSng(FenetreMatrice.GridPmat.Text)
               End If
            Next j%
         Next i%
         ' *************************************************
         '-----------------------------------------------------
         ' Création du fichier d'éléments de matrice
         ' et écriture de ces éléments dans le fichier
         '-----------------------------------------------------
         Open Maths.ctrlCMDialog.FileName For Output As #1
            Write #1, OrdreMat%
            For i% = 1 To OrdreMat%
               For j% = 1 To OrdreMat%
                  Write #1, Pmat(i%, j%)
               Next j%
            Next i%
         Close #1
         ' *************************************************
      Else
         Message$ = " La matrice produit P n'a pas été calculée ! "
         MsgBox Message$, 16, "EnregMat"
      End If
   Case "Vap"
      If lblWmat.Caption = "Valeurs propres :" Then
         ' *************************************************
         Maths.ctrlCMDialog.DefaultExt = "vap"
         Maths.ctrlCMDialog.Filter = "Valeurs propres (*.vap)|*.vap"
         Maths.ctrlCMDialog.Flags = &H2&
         Maths.ctrlCMDialog.Action = 2
         ' *************************************************
         If lblPmat.Caption = "Valeurs propres :" Then
            ' certaines valeurs propres complexes
            '-----------------------------------------------------
            ' Création du fichier des valeurs propres
            ' et écriture de ces éléments dans le fichier
            '-----------------------------------------------------
            Open Maths.ctrlCMDialog.FileName For Output As #1
            Write #1, OrdreMat%
            ' Indication du caractère complexe
            ' de certaines valeurs propres
            Write #1, "1"
            GridWmat.Col = 1
            GridPmat.Col = 1
            For i% = 1 To OrdreMat%
               GridWmat.Row = i%
               GridPmat.Row = i%
               Write #1, GridWmat.Text
               Write #1, GridPmat.Text
               Next i%
            Close #1
            ' *************************************************
         Else
            ' toutes les valeurs propres réelles
            '-----------------------------------------------------
            ' Création du fichier des valeurs propres
            ' et écriture de ces éléments dans le fichier
            '-----------------------------------------------------
            Open Maths.ctrlCMDialog.FileName For Output As #1
               Write #1, OrdreMat%
               ' Indication du caractère réel
               ' de toutes les valeurs propres
               Write #1, "0"
               GridWmat.Col = 1
               For i% = 1 To OrdreMat%
                  GridWmat.Row = i%
                  Write #1, GridWmat.Text
               Next i%
            Close #1
            ' *************************************************
         End If
      Else
         Message$ = " Les valeurs propres de M n'ont pas été calculées ! "
         MsgBox Message$, 16, "EnregMat"
      End If
   Case "Vep"
      ' *************************************************
      Maths.ctrlCMDialog.DefaultExt = "vep"
      Maths.ctrlCMDialog.Filter = "Vecteurs propres (*.vep)|*.vep"
      Maths.ctrlCMDialog.Flags = &H2&
      Maths.ctrlCMDialog.Action = 2
      ' *************************************************
      If lblPRmat.Caption = "Matrice des vecteurs propres :" Then
         ' *************************************************
         ' tous les vecteurs propres réels
         '-----------------------------------------------------
         ' Création du fichier des vecteurs propres
         ' et écriture de ces éléments dans le fichier
         '-----------------------------------------------------
         Open Maths.ctrlCMDialog.FileName For Output As #1
         Write #1, OrdreMat%
         ' Indication du caractère réel
         ' de tous les vecteurs propres
         Write #1, "0"
         For i% = 1 To OrdreMat%
            For j% = 1 To OrdreMat%
               Write #1, Pmat(i%, j%)
            Next j%
         Next i%
         Close #1
         ' *************************************************
      ElseIf lblPRmat.Caption = "Matrice des vecteurs propres (partie réelle):" Then
         ' certains vecteurs propres complexes
         '-----------------------------------------------------
         ' Création du fichier des vecteurs propres
         ' et écriture de ces éléments dans le fichier
         '-----------------------------------------------------
         Open Maths.ctrlCMDialog.FileName For Output As #1
         Write #1, OrdreMat%
         ' Indication du caractère complexe
         ' de certains vecteurs propres
         Write #1, "1"
         For i% = 1 To OrdreMat%
            For j% = 1 To OrdreMat%
               Write #1, Pmat(i%, j%)
            Next j%
         Next i%
         Close #1
         ' *************************************************
      Else
         Message$ = " Les vecteurs propres de M n'ont pas été calculés ! "
         MsgBox Message$, 16, "EnregMat"
      End If
   Case "W"
      If lblWmat.Caption = "Matrice W inverse de M :" Then
         ' *************************************************
         Maths.ctrlCMDialog.DefaultExt = "mat"
         Maths.ctrlCMDialog.Filter = "Matrice (*.mat)|*.mat"
         Maths.ctrlCMDialog.Flags = &H2&
         Maths.ctrlCMDialog.Action = 2
         ' ********************************************
         '               ReDim
         ' ********************************************
         Erase Wmat
         ReDim Wmat(1 To OrdreMat%, 1 To OrdreMat%)
         ' *******************************************************
         ' affectation de leurs valeurs aux éléments de la matrice
         ' *******************************************************
         For i% = 1 To OrdreMat%
            For j% = 1 To OrdreMat%
               FenetreMatrice.GridWmat.Row = i%
               FenetreMatrice.GridWmat.Col = j%
               If FenetreMatrice.GridWmat.Text = "" Then
                  FenetreMatrice.GridWmat.Text = 0
                  Wmat(i%, j%) = 0
               Else
                  Wmat(i%, j%) = CSng(FenetreMatrice.GridWmat.Text)
               End If
            Next j%
         Next i%
         ' *************************************************
         '-----------------------------------------------------
         ' Création du fichier d'éléments de matrice
         ' et écriture de ces éléments dans le fichier
         '-----------------------------------------------------
         Open Maths.ctrlCMDialog.FileName For Output As #1
            Write #1, OrdreMat%
            For i% = 1 To OrdreMat%
               For j% = 1 To OrdreMat%
                  Write #1, Wmat(i%, j%)
               Next j%
            Next i%
         Close #1
         ' *************************************************
      Else
         Message$ = " La matrice W inverse de M n'a pas été calculée ! "
         MsgBox Message$, 16, "EnregMat"
      End If
   End Select
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnregMat:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub
Public Sub DiaMat()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Mmat, Pmat, RacineR, RacineI
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Pmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim RacineR(1 To OrdreMat%)
   ReDim RacineI(1 To OrdreMat%)
   ' *************************************************
   '               Initialisations
   ' *************************************************
   lblValDet.Caption = ""
   lblInfo.ForeColor = ROUGE
   lblInfo.Font.bold = True
   lblInfo.Font.underline = False
   lblInfo.Caption = "CALCUL EN COURS..."
   'DoEvents
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Mmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridMmat.Row = i%
         FenetreMatrice.GridMmat.Col = j%
         If FenetreMatrice.GridMmat.Text = "" Then
            FenetreMatrice.GridMmat.Text = 0
            Mmat(i%, j%) = 0
         Else
            Mmat(i%, j%) = CSng(FenetreMatrice.GridMmat.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   ' calcul des valeurs propres et vecteurs propres
   ' *************************************************
   Select Case NumMethodDiag%
   Case 1
      Call DiaMaQR
      ' Méthode du double QR
   Case 2
      Call DiaMaSoBaSS
      ' SOURIAU+BAIRSTOW+SystèmeSurdéterminé
   Case 3
      Call VerifieMatSym(OrdreMat%, Mmat())
      If Erreur = True Then
         lblInfo.Caption = ""
         Exit Sub
      End If
      racinecomplexe% = 0
      Call DiaMaJacobi(OrdreMat%, Mmat(), RacineR(), Pmat())
   End Select
   ' *************************************************
   ' affichage de la partie réelle des valeurs propres
   ' sur une colonne
   ' *************************************************
   ' 1)- mise en place de la grille gridWmat
   ' **************************************
   GridWmat.Rows = OrdreMat% + 1
   GridWmat.Cols = 2
   '***************************************
   GridWmat.FixedAlignment(0) = 2
   GridWmat.ColWidth(0) = 500
   GridWmat.FixedAlignment(1) = 2
   GridWmat.ColWidth(1) = 3000
   ' ********************************
   ' numérotation 1ère ligne
   ' ********************************
   GridWmat.Row = 0
   GridWmat.Col = 1
   If racinecomplexe% = 0 Then
      GridWmat.Text = ""
   Else
      GridWmat.Text = "partie réelle"
   End If
   ' *********************************
   ' numérotation 1ère colonne
   ' *********************************
   GridWmat.Col = 0
   For i% = 1 To OrdreMat%
      GridWmat.Row = i%
      GridWmat.Text = Format(i%, "0")
   Next i%
   ' *****************************************************
   ' 2)- affichage de la partie réelle des valeurs propres
   ' *****************************************************
   lblWmat.Caption = "Valeurs propres :"
   For i% = 1 To OrdreMat%
      GridWmat.Row = i%
      GridWmat.Col = 1
      GridWmat.Text = ""
      GridWmat.Text = Format(RacineR(i%), "0.000")
   Next i%
   ' **************************************************************
   ' affichage éventuel de la partie imaginaire des valeurs propres
   ' sur une colonne
   ' **************************************************************
   ' 1)- mise en place de la grille gridPmat
   ' **************************************
   GridPmat.Rows = OrdreMat% + 1
   GridPmat.Cols = 2
   '***************************************
   GridPmat.FixedAlignment(0) = 2
   GridPmat.ColWidth(0) = 500
   GridPmat.FixedAlignment(1) = 2
   GridPmat.ColWidth(1) = 3000
   ' ********************************
   ' numérotation 1ère ligne
   ' ********************************
   GridPmat.Row = 0
   GridPmat.Col = 1
   If racinecomplexe% = 0 Then
      GridPmat.Text = ""
   Else
      GridPmat.Text = "partie imaginaire"
   End If
   ' *********************************
   ' numérotation 1ère colonne
   ' *********************************
   GridPmat.Col = 0
   For i% = 1 To OrdreMat%
      GridPmat.Row = i%
      GridPmat.Text = Format(i%, "0")
   Next i%
   ' *********************************************************
   ' 2)- affichage de la partie imaginaire des valeurs propres
   ' *********************************************************
   If racinecomplexe% = 0 Then
      lblPmat.Caption = "VALEURS PROPRES TOUTES REELLES"
      For i% = 1 To OrdreMat%
         GridPmat.Row = i%
         GridPmat.Col = 1
         GridPmat.Text = ""
      Next i%
   Else
      lblPmat.Caption = "Valeurs propres :"
      For i% = 1 To OrdreMat%
         GridPmat.Row = i%
         GridPmat.Col = 1
         GridPmat.Text = ""
         GridPmat.Text = Format(RacineI(i%), "0.000")
      Next i%
   End If
   ' *************************************************
   ' affichage des vecteurs propres
   ' *************************************************
   If racinecomplexe% = 0 Then
      ' ***************************************
      ' Toutes les valeurs propres sont réelles
      ' ***************************************
      ' *****************************************************
      ' affichage des éléments
      ' *****************************************************
      lblPRmat.Caption = "Matrice des vecteurs propres :"
      For i% = 1 To OrdreMat%
         FenetreMatrice.GridPRmat.Row = i%
         For j% = 1 To OrdreMat%
            GridPRmat.Col = j%
            GridPRmat.Text = ""
            GridPRmat.Text = Format(Pmat(i%, j%), "0.000")
         Next j%
      Next i%
      ' *****************************************************
      ' affichage des éléments
      ' *****************************************************
      lblPImat.Caption = ""
      For i% = 1 To OrdreMat%
         GridPImat.Row = i%
         For j% = 1 To OrdreMat%
            GridPImat.Col = j%
            GridPImat.Text = ""
         Next j%
      Next i%
      ' *****************************************************
   Else
      ' Certaines valeurs propres sont complexes
      If NumMethodDiag% = 1 Then
         ' *********************************************
         ' Les vecteurs propres sont calculés, y compris
         ' les vecteurs propres complexes correspondant
         ' aux valeurs propres complexes
         ' *********************************************
         ' affichage des éléments
         ' *****************************************************
         lblPRmat.Caption = "Matrice des vecteurs propres (partie réelle):"
         lblPImat.Caption = "Matrice des vecteurs propres (partie imaginaire):"
         j% = 0
         Do
            j% = j% + 1
            If RacineI(j%) = 0 Then
               GridPRmat.Col = j%
               GridPImat.Col = j%
               For i% = 1 To OrdreMat%
                  GridPRmat.Row = i%
                  GridPRmat.Text = ""
                  GridPRmat.Text = Format(Pmat(i%, j%), "0.000")
                  GridPImat.Row = i%
                  GridPImat.Text = "0"
               Next i%
            Else
               GridPRmat.Col = j%
               GridPImat.Col = j%
               For i% = 1 To OrdreMat%
                  GridPRmat.Row = i%
                  GridPRmat.Text = ""
                  GridPRmat.Text = Format(Pmat(i%, j%), "0.000")
                  GridPImat.Row = i%
                  GridPImat.Text = ""
                  GridPImat.Text = Format(Pmat(i%, j% + 1), "0.000")
               Next i%
               GridPRmat.Col = j% + 1
               GridPImat.Col = j% + 1
               For i% = 1 To OrdreMat%
                  GridPRmat.Row = i%
                  GridPRmat.Text = ""
                  GridPRmat.Text = Format(Pmat(i%, j%), "0.000")
                  GridPImat.Row = i%
                  GridPImat.Text = ""
                  GridPImat.Text = Format(-Pmat(i%, j% + 1), "0.000")
               Next i%
               j% = j% + 1
            End If
         Loop Until j% = OrdreMat%
      Else
         ' *********************************************
         ' Les vecteurs propres ne peuvent être calculés.
         ' *********************************************
         Message$ = " Certaines valeurs propres sont complexes "
         Message$ = Message$ & Chr$(13) & " Les vecteurs propres ne seront pas calculés."
         MsgBox Message$, 16, "DiaMat"
      End If
   End If
   '***************************************
   '************ Explications *************
   '***************************************
   Info$ = "Dernier calcul effectué : Diagonalisation de M : "
   Select Case NumMethodDiag%
   Case 1
      Info$ = Info$ & "Méthode n°1 : Méthode du double QR avec déplacement."
   Case 2
      Info$ = Info$ & "Méthode n°2 : Détermination du polynôme caractéristique (méthode de SOURIAU), recherche de ses zéros "
      Info$ = Info$ & "(méthode de BAIRSTOW) et détermination des vecteurs propres par résolution de systèmes sur-déterminés."
   Case 3
      Info$ = Info$ & "Méthode n°3 : Méthode de JACOBI classique (pour une matrice symétrique)."
   End Select
   lblInfo.ForeColor = BLEU
   lblInfo.Font.bold = True
   lblInfo.Font.underline = False
   lblInfo.Caption = Info$
   ' *************************************************
End Sub
Public Sub DiaMaSoBaSS()
' -------------------------------------------------------------
' David BLUM
' 15/01/1997
'    Diagonalisation de matrice (méthode n°1) :
' 1] Détermination du polynôme caractéristique
' (méthode de SOURIAU);
' 2] Recherche des zéros de ce polynôme
' (méthode de BAIRSTOW);
' 3] Détermination des vecteurs propres
' par résolution de systèmes sur-déterminés."
' ------------------------------------------------
'1] Détermination du polynôme caractéristique
' (méthode de SOURIAU)
' ------------------------------------------------
' (-1)^n
mupun% = (-1) ^ OrdreMat%
' ----------------------------------------------------------------------
' Dims
ReDim MA(1 To OrdreMat%, 1 To OrdreMat%)      ' matrice intermédiaire de calcul
ReDim MB(1 To OrdreMat%, 1 To OrdreMat%)      ' matrice intermédiaire de calcul
ReDim Ppol(0 To OrdreMat%)             ' coefficients du polynôme caractéristique
ReDim apol(0 To OrdreMat%)             ' coefficients intermédiaires de calcul
' ----------------------------------------------------------------------
' Initialisation de MA et MB
For iloc% = 1 To OrdreMat%
   For jloc% = 1 To OrdreMat%
      MA(iloc%, jloc%) = 0
      MB(iloc%, jloc%) = 0
   Next jloc%
Next iloc%
' -------------------------------------------
' Initialisations
apol(0) = 1
' -------------------------------------------
' Itérations
For i% = 1 To OrdreMat%
   ' ----------------------------------------------------------------
   For iloc% = 1 To OrdreMat%
      MA(iloc%, iloc%) = MA(iloc%, iloc%) + apol(i% - 1)
   Next iloc%
   For iloc% = 1 To OrdreMat%
      For jloc% = 1 To OrdreMat%
         For kloc% = 1 To OrdreMat%
            MB(iloc%, jloc%) = MB(iloc%, jloc%) + Mmat(iloc%, kloc%) * MA(kloc%, jloc%)
         Next kloc%
      Next jloc%
   Next iloc%
   For iloc% = 1 To OrdreMat%
      For jloc% = 1 To OrdreMat%
         MA(iloc%, jloc%) = MB(iloc%, jloc%)
         MB(iloc%, jloc%) = 0
      Next jloc%
   Next iloc%
   ' -------------------
   ' Trace de MA
   Trace = 0
   For iloc% = 1 To OrdreMat%
      Trace = Trace + MA(iloc%, iloc%)
   Next iloc%
   ' -------------------
   apol(i%) = -Trace / i%
   ' ----------------------------------------------------------------
Next i%
' -----------------------------------------------------------------
' Polynôme caractéristique de Mmat :
For i% = 0 To OrdreMat%
   Ppol(i%) = mupun% * apol(OrdreMat% - i%)
Next i%
DegPpol% = OrdreMat%
' -----------------------------------------------------------------
' RECHERCHE DES ZEROS DU POLYNOME CARACTERISTIQUE
' PAR LA METHODE DE BAIRSTOW
' -----------------------------------------------------------------
Call ZerosPolBairstow
' -----------------------------------------------------------------
' S'il existe des racines complexes,
' ne pas chercher à calculer les vecteurs propres
' -----------------------------------------------------------------
If NbComplexe% = 1 Then
   Message$ = " Certaines valeurs propres sont complexes "
   Message$ = Message$ & Chr$(13) & " Les vecteurs propres ne seront pas calculés."
   MsgBox Message$, 16, "DiaMat01"
   Exit Sub
End If
' ----------------------------------------------------------------------
' RECHERCHE DES VECTEURS PROPRES
' PAR RESOLUTION DE SYSTEMES SURDETERMINES
' -------------------------------------------------------------------------
OrdMatMoinsUn% = OrdreMat% - 1
' Dims
ReDim mp(1 To OrdreMat%, 1 To OrdreMat%)             ' matrice de passage
                                                   ' = matrice des vecteurs propres (en colonne)
ReDim MLI(1 To OrdreMat%, 1 To OrdreMat%)            ' matrice (Mat - Lambda * I)
ReDim MLIC(1 To OrdreMat%, 1 To OrdMatMoinsUn%)      ' matrice (MLI - colonne numcolonne% de MLI)
ReDim MMC(1 To OrdMatMoinsUn%, 1 To OrdMatMoinsUn%)  ' matrice symétrique ((transposée de MLIC) * MLIC)
ReDim WWC(1 To OrdMatMoinsUn%, 1 To OrdMatMoinsUn%)  ' matrice inverse de MMC
ReDim DM(1 To OrdMatMoinsUn%)                        ' vecteur deuxième membre
                                                     ' -(transposée de MLIC) * (colonne numcolonne% de MLI)
' -------------------------------------------------------------------------
' Principe du calcul :
' (colonne numvap% de MP) = WWC * DM
' -------------------------------------------------------------------------
' pour chaque valeur propre...
For numvap% = 1 To OrdreMat%
   ' *******************
   ' construction de MLI
   ' *******************
   For ligne% = 1 To OrdreMat%
      For colonne% = 1 To OrdreMat%
         If ligne% = colonne% Then
            MLI(ligne%, colonne%) = Mmat(ligne%, colonne%) - RacineR(numvap%)
         Else
            MLI(ligne%, colonne%) = Mmat(ligne%, colonne%)
         End If
      Next colonne%
   Next ligne%
   For numcolonne% = OrdreMat% To 1 Step -1
      ' ********************
      ' construction de MLIC
      ' ********************
      For colonne% = 1 To OrdMatMoinsUn%
         If colonne% < numcolonne% Then
            For ligne% = 1 To OrdreMat%
               MLIC(ligne%, colonne%) = MLI(ligne%, colonne%)
            Next ligne%
         Else
            For ligne% = 1 To OrdreMat%
               MLIC(ligne%, colonne%) = MLI(ligne%, colonne% + 1)
            Next ligne%
         End If
      Next colonne%
      ' ********************************
      ' construction de MMC (symétrique)
      ' ********************************
      For ligne% = 1 To OrdMatMoinsUn%
         For colonne% = 1 To OrdMatMoinsUn%
            MMC(ligne%, colonne%) = 0
         Next colonne%
      Next ligne%
      For ligne% = 1 To OrdMatMoinsUn%
         For colonne% = 1 To OrdMatMoinsUn%
            For iloc% = 1 To OrdreMat%
               MMC(ligne%, colonne%) = MMC(ligne%, colonne%) + MLIC(iloc%, ligne%) * MLIC(iloc%, colonne%)
            Next iloc%
         Next colonne%
      Next ligne%
      ' ************************************************
      ' inversion de MMC
      ' ************************************************
      DePo = True
      Call InvMatCholeski(OrdMatMoinsUn%, DePo, DetMMC, MMC(), WWC())
      ' ********** fin d'inversion de MMC **************
      If DePo = True Then
      ' MMC est inversible
         ' ******************************************************
         ' calcul du deuxième membre
         ' DM = - (transposée de MLIC) * (colonne numcolonne% de MLI)
         ' ******************************************************
         For iloc% = 1 To OrdMatMoinsUn%
            DM(iloc%) = 0
            For jloc% = 1 To OrdreMat%
               DM(iloc%) = DM(iloc%) - MLIC(jloc%, iloc%) * MLI(jloc%, numcolonne%)
            Next jloc%
         Next iloc%
         ' *******************************************
         ' calcul du vecteur propre MP(iloc%, numvap%)
         ' *******************************************
         difind% = 0
         For iloc% = 1 To OrdreMat%
            If iloc% = numcolonne% Then
               mp(iloc%, numvap%) = 1
               difind% = 1
            Else
               mp(iloc%, numvap%) = 0
               For jloc% = 1 To OrdMatMoinsUn%
                  mp(iloc%, numvap%) = mp(iloc%, numvap%) + WWC(iloc% - difind%, jloc%) * DM(jloc%)
               Next jloc%
            End If
         Next iloc%
         Exit For
      End If
      ' DePo = False
      ' MMC n'est pas inversible
      ' la composante numcolonne% du vecteur propre est sans doute nulle
      ' on va essayer avec une autre composante,
      ' sauf si on les a déjà toutes essayées, auquel cas :
      If numcolonne% = 1 Then
         Message$ = " Echec dans la recherches des vecteurs propres."
         MsgBox Message$, 16, "DiaMat01"
         Exit Sub
      End If
   Next numcolonne%
Next numvap%
' --------------------------------------------------
' Normalisation des vecteurs propres
' et remplissage des matrices avec les résultats
' ( Pmat =  Matrice de passage                     )
' (      =  matrice des vecteurs propres en colonne)
' ----------------------------------
For jloc% = 1 To OrdreMat%
   elem = 0
   For iloc% = 1 To OrdreMat%
      elem = elem + mp(iloc%, jloc%) ^ 2
   Next iloc%
   elem = Sqr(elem)
   For iloc% = 1 To OrdreMat%
      Pmat(iloc%, jloc%) = mp(iloc%, jloc%) / elem
   Next iloc%
Next jloc%
' ---------------------------------------------------------
End Sub
Public Sub DiaMaQR2Vap(OrdreDqr%, NbComplexe%, Gdqr(), VaPRdqr(), VaPIdqr())
' **************************************************
' Recherche des valeurs propres réelles et complexes
' d'une matrice réelle par l'algorithme du double QR
' **************************************************
' En entrée :
' OrdreDqr%    = ordre des matrices en jeu
' Gdqr()       = matrice de départ M déjà sous forme de la
'                matrice de HESSENBERG Gdqr
' En sortie :
' NbComplexe%  = nombre de valeurs propres complexes
' VaPRdqr()    = partie réelle des valeurs propres
' VaPIdqr()    = partie imaginaire des valeurs propres
' -------------------------------------------------------------------
eps = 0.000001
' ---------------------------------------------------
' Examen de OrdreDqr%
' ---------------------------------------------------
If OrdreDqr% < 3 Then
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "DiaMaQR2Vap"
End If
' ---------------------------------------------------
' DIMs
' ---------------------------------------------------
ReDim GKdqr(1 To OrdreDqr%, 1 To OrdreDqr%)  ' matrice de Hessenberg supérieure
'                                            ' (matrice de travail)
ReDim VaPRdqr(1 To OrdreDqr%)                ' partie réelle des valeurs propres
ReDim VaPIdqr(1 To OrdreDqr%)                ' partie imaginaire des valeurs propres
' *****************************************************
' On suppose que M a été mise sous forme
' d'une matrice de HESSENBERG G (par la Sub Hessenberg)
' *****************************************************
' Algorithme du double QR avec translation
' ****************************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
' On met Gdqr dans la matrice de travail GKdqr
' pour ne pas la modifier en sortie
For i% = 1 To OrdreDqr%
   For j% = 1 To OrdreDqr%
      GKdqr(i%, j%) = Gdqr(i%, j%)
   Next j%
Next i%
' ---------------------------------------------------
ierr% = 0
norm = 0
k% = 1
NbComplexe% = 0
' ---------------------------------------------------
' Calcul de la norme de la matrice
' ---------------------------------------------------
For i% = 1 To OrdreDqr%
   For j% = k% To OrdreDqr%
      norm = norm + Abs(GKdqr(i%, j%))
   Next j%
   k% = i%
Next i%
' ---------------------------------------------------
en% = OrdreDqr%
itn% = 30 * OrdreDqr%
Tloc = 0
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' Recherche de chaque valeur propre
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Do
   If en% < 1 Then
      Exit Sub
   End If
   its% = 0
   na% = en% - 1
   enm2% = na% - 1
   ' ---------------------------------------------------
   ' Recherche un petit élément sous-diagonal isolé
   ' pour ll% = en% pas -1 jusqu'à 1
   ' ---------------------------------------------------
   Do
      For ll% = 1 To en%
         l% = en% + 1 - ll%
         If l% = 1 Then
            Exit For
         End If
         Sloc = Abs(GKdqr(l% - 1, l% - 1)) + Abs(GKdqr(l%, l%))
         If Sloc = 0 Then
            Sloc = norm
         End If
         tst1 = Sloc
         tst2 = tst1 + Abs(GKdqr(l%, l% - 1))
         If tst2 = tst1 Then
            Exit For
         End If
      Next ll%
      ' ---------------------------------------------------
      ' forme un déplacement ("shift")
      ' ---------------------------------------------------
      Xloc = GKdqr(en%, en%)
      If l% = en% Then
         Exit Do
      End If
      Yloc = GKdqr(na%, na%)
      Wloc = GKdqr(en%, na%) * GKdqr(na%, en%)
      If l% = na% Then
         Exit Do
      End If
      If itn% = 0 Then
         Erreur = True
         Message$ = " Pas de convergence après " & 30 * OrdreDqr% & " itérations" & Chr$(13)
         Message$ = Message$ & " pour la valeur propre numéro " & en%
         MsgBox Message$, 16, "DiaMaQR2Vap"
         Exit Sub
      End If
      If its% = 10 Or its% = 20 Then
         ' ---------------------------------------------------
         ' forme un déplacement exceptionnel
         ' ---------------------------------------------------
         Tloc = Tloc + Xloc
         For i% = 1 To en%
            GKdqr(i%, i%) = GKdqr(i%, i%) - Xloc
         Next i%
         Sloc = Abs(GKdqr(en%, na%)) + Abs(GKdqr(na%, enm2%))
         Xloc = 0.75 * Sloc
         Yloc = Xloc
         Wloc = -0.4375 * Sloc * Sloc
      End If
      ' ---------------------------------------------------
      its% = its% + 1
      itn% = itn% - 1
      ' ---------------------------------------------------
      ' Recherche deux petits éléments sous-diagonaux
      ' consécutifs pour m% = en%-2 pas -1 jusqu'à l%
      ' ---------------------------------------------------
      For mm% = l% To enm2%
         mloc% = enm2% + l% - mm%
         Zloc = GKdqr(mloc%, mloc%)
         Rloc = Xloc - Zloc
         Sloc = Yloc - Zloc
         Ploc = (Rloc * Sloc - Wloc) / GKdqr(mloc% + 1, mloc%) + GKdqr(mloc%, mloc% + 1)
         Qloc = GKdqr(mloc% + 1, mloc% + 1) - Zloc - Rloc - Sloc
         Rloc = GKdqr(mloc% + 2, mloc% + 1)
         Sloc = Abs(Ploc) + Abs(Qloc) + Abs(Rloc)
         Ploc = Ploc / Sloc
         Qloc = Qloc / Sloc
         Rloc = Rloc / Sloc
         If mloc% = l% Then
            Exit For
         End If
         tst1 = Abs(Ploc) * (Abs(GKdqr(mloc% - 1, mloc% - 1)) + Abs(Zloc) + Abs(GKdqr(mloc% + 1, mloc% + 1)))
         tst2 = tst1 + Abs(GKdqr(mloc%, mloc% - 1)) * (Abs(Qloc) + Abs(Rloc))
         If tst2 = tst1 Then
            Exit For
         End If
      Next mm%
      mp2% = mloc% + 2
      For i% = mp2% To en%
         GKdqr(i%, i% - 2) = 0
         If i% <> mp2% Then
            GKdqr(i%, i% - 3) = 0
         End If
      Next i%
      ' ---------------------------------------------------
      ' Etape de double QR impliquant les lignes l% à en%
      ' et les colonnes m% à en%
      ' ---------------------------------------------------
      For k% = mloc% To na%
         If k% <> mloc% Then
            Ploc = GKdqr(k%, k% - 1)
            Qloc = GKdqr(k% + 1, k% - 1)
            Rloc = 0
            If k% <> na% Then
               Rloc = GKdqr(k% + 2, k% - 1)
            End If
            Xloc = Abs(Ploc) + Abs(Qloc) + Abs(Rloc)
            If Xloc <> 0 Then
               Ploc = Ploc / Xloc
               Qloc = Qloc / Xloc
               Rloc = Rloc / Xloc
            End If
         End If
         If Xloc <> 0 Then
            Sloc = Sgn(Ploc) * Sqr(Ploc * Ploc + Qloc * Qloc + Rloc * Rloc)
            If k% = mloc% Then
               If l% <> mloc% Then
                  GKdqr(k%, k% - 1) = -GKdqr(k%, k% - 1)
               End If
            Else
               GKdqr(k%, k% - 1) = -Sloc * Xloc
            End If
            Ploc = Ploc + Sloc
            Xloc = Ploc / Sloc
            Yloc = Qloc / Sloc
            Zloc = Rloc / Sloc
            Qloc = Qloc / Ploc
            Rloc = Rloc / Ploc
            If k% = na% Then
               ' ---------------------------------------------------
               ' Modification de ligne
               ' ---------------------------------------------------
               For j% = k% To en%
                  Ploc = GKdqr(k%, j%) + Qloc * GKdqr(k% + 1, j%)
                  GKdqr(k%, j%) = GKdqr(k%, j%) - Ploc * Xloc
                  GKdqr(k% + 1, j%) = GKdqr(k% + 1, j%) - Ploc * Yloc
               Next j%
               If en% < k% + 3 Then
                  j% = en%
               Else
                  j% = k% + 3
               End If
               ' ---------------------------------------------------
               ' Modification de colonne
               ' ---------------------------------------------------
               For i% = l% To j%
                  Ploc = Xloc * GKdqr(i%, k%) + Yloc * GKdqr(i%, k% + 1)
                  GKdqr(i%, k%) = GKdqr(i%, k%) - Ploc
                  GKdqr(i%, k% + 1) = GKdqr(i%, k% + 1) - Ploc * Qloc
               Next i%
               ' ---------------------------------------------------
            Else
               ' ---------------------------------------------------
               ' Modification de ligne
               ' ---------------------------------------------------
               For j% = k% To en%
                  Ploc = GKdqr(k%, j%) + Qloc * GKdqr(k% + 1, j%) + Rloc * GKdqr(k% + 2, j%)
                  GKdqr(k%, j%) = GKdqr(k%, j%) - Ploc * Xloc
                  GKdqr(k% + 1, j%) = GKdqr(k% + 1, j%) - Ploc * Yloc
                  GKdqr(k% + 2, j%) = GKdqr(k% + 2, j%) - Ploc * Zloc
               Next j%
               If en% < k% + 3 Then
                  j% = en%
               Else
                  j% = k% + 3
               End If
               ' ---------------------------------------------------
               ' Modification de colonne
               ' ---------------------------------------------------
               For i% = l% To j%
                  Ploc = Xloc * GKdqr(i%, k%) + Yloc * GKdqr(i%, k% + 1) + Zloc * GKdqr(i%, k% + 2)
                  GKdqr(i%, k%) = GKdqr(i%, k%) - Ploc
                  GKdqr(i%, k% + 1) = GKdqr(i%, k% + 1) - Ploc * Qloc
                  GKdqr(i%, k% + 2) = GKdqr(i%, k% + 2) - Ploc * Rloc
               Next i%
               ' ---------------------------------------------------
            End If
         End If
      Next k%
   Loop
   If l% = en% Then
      ' ---------------------------------------------------
      ' Une racine trouvée
      ' ---------------------------------------------------
      VaPRdqr(en%) = Xloc + Tloc
      VaPIdqr(en%) = 0
      en% = na%
   ElseIf l% = na% Then
      ' ---------------------------------------------------
      ' Deux racines trouvées
      ' ---------------------------------------------------
      Ploc = (Yloc - Xloc) / 2
      Qloc = Ploc * Ploc + Wloc
      Zloc = Sqr(Abs(Qloc))
      Xloc = Xloc + Tloc
      If Qloc >= 0 Then
         ' ---------------------------------------------------
         ' Paire réelle
         ' ---------------------------------------------------
         Zloc = Ploc + Sgn(Ploc) * Zloc
         VaPRdqr(na%) = Xloc + Zloc
         VaPRdqr(en%) = Xloc + Zloc
         If Zloc <> 0 Then
            VaPRdqr(en%) = Xloc - Wloc / Zloc
         End If
         VaPIdqr(na%) = 0
         VaPIdqr(en%) = 0
      Else
         ' ---------------------------------------------------
         ' Paire complexe
         ' ---------------------------------------------------
         NbComplexe% = NbComplexe% + 2
         VaPRdqr(na%) = Xloc + Ploc
         VaPRdqr(en%) = Xloc + Ploc
         VaPIdqr(na%) = Zloc
         VaPIdqr(en%) = -Zloc
      End If
         ' ---------------------------------------------------
      en% = enm2%
   End If
Loop
End Sub
Public Sub DiaMaQR()
' ******************************************************
' Recherche des éléments propres d'une matrice réelle M;
' 1) Transformation de M en une matrice de type
' HESSENBERG supérieure G;
' 2) Recherche de valeurs et vecteurs propres de G
' par l'algorithme du double QR avec déplacement.
' *********************************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
eps = 0.000001
Erreur = False
' ---------------------------------------------------
' Examen de OrdreMat%
' ---------------------------------------------------
If OrdreMat% < 3 Then
   Erreur = True
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "DiaMaQR"
   Exit Sub
End If
' ---------------------------------------------------
' DIMs
' ---------------------------------------------------
ReDim Gmat(1 To OrdreMat%, 1 To OrdreMat%)   ' matrice de Hessenberg supérieure
ReDim PMGmat(1 To OrdreMat%, 1 To OrdreMat%) ' matrice de passage de M à G
ReDim PermVec%(1 To OrdreMat%)                ' vecteur permutation M -> G
' ***********************************
' On commence par mettre M sous forme
' d'une matrice de HESSENBERG G
' ***********************************
Call Hessenberg(OrdreMat%, Mmat(), Gmat(), PermVec%())
Call PassMatHes(OrdreMat%, PMGmat(), Gmat(), PermVec%())
Erase PermVec%
' +++++++++++++++++++++++++++++++++++++++++++++
' A ce stade :
' G = (transposée de PMG) x M x PMG
' ****************************************
' Algorithme du double QR avec translation
' ****************************************
Call DiaMaQR2VapVep(OrdreMat%, racinecomplexe%, Gmat(), PMGmat(), Pmat(), RacineR(), RacineI())
If Erreur = True Then Exit Sub
' ---------------------------------------------------
End Sub
Public Sub DiaMaJacobi(OrdreJa%, Mja(), Vja(), Pja())
' *********************************************
'  Recherche des éléments propres d'une matrice
'  symétrique réelle par la méthode de JACOBI
'  classique
' *********************************************
' En entrée :
' OrdreJa%     = ordre des matrices en jeu
' Mja()        = matrice symétrique à diagonaliser
' En sortie :
' Pja()        = matrice unitaire de passage de M à Wqr
'               (vecteurs propres en colonnes)
' Vja()        = valeurs propres
' ---------------------------------------------------
' Examen de OrdreJa%
' ---------------------------------------------------
If OrdreJa% < 3 Then
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "DiaJacobi"
End If
' ---------------------------------------------------
' DIMs
' -------------------------------------------------------------------------
ReDim MKja(1 To OrdreJa%, 1 To OrdreJa%)     ' matrice de travail
ReDim Qja(1 To OrdreJa%, 1 To OrdreJa%)      ' matrice unitaire
ReDim QINVja(1 To OrdreJa%, 1 To OrdreJa%)   ' matrice inverse de Qja
ReDim Wja(1 To OrdreJa%, 1 To OrdreJa%)      ' matrice diagonale de travail
ReDim Vja0(1 To OrdreJa%)                  ' valeur propres de travail
' -------------------------------------------------------------------------
' Initialisation de Wja et Pja
' ---------------------------------------------------------
For i% = 1 To OrdreJa%
   For j% = i% To OrdreJa%
      Wja(i%, j%) = Mja(i%, j%)
      If i% = j% Then
         Pja(i%, j%) = 1
      Else
         Pja(i%, j%) = 0
         Pja(j%, i%) = 0
         Wja(j%, i%) = Wja(i%, j%)
      End If
   Next j%
Next i%
' ---------------------------------------------------------
' Précision de convergence
eps = 0.01
' ---------------------------------------------------------
' Nombre de cycles à effectuer
numcyclesmax% = 100
' ---------------------------------------------------------
' Début du cycle numcycle%
numcycle% = 0
Do
   numcycle% = numcycle% + 1
   ' ----------------------------
   ' tolérance dynamique
   epsdyn = 1 / (100 ^ numcycle%)
   ' ----------------------------
   For i% = 1 To OrdreJa% - 1
      For j% = i% + 1 To OrdreJa%
         ' ----------------------------------------------------------------------
         ' facteur de couplage
         If Wja(i%, i%) = 0 Or Wja(j%, j%) = 0 Then
            factcoupl = 10 * epsdyn
         Else
            factcoupl = Abs(Wja(i%, j%)) / Sqr(Abs(Wja(i%, i%) * Wja(j%, j%)))
         End If
         ' ----------------------------------------------------------------------
         If factcoupl > epsdyn Then
            If Abs(Wja(i%, j%)) < epsdyn Then
               If Wja(i%, i%) = Wja(j%, j%) Then
                  Message$ = "Diagonalisation impossible"
                  MsgBox Message$, 16, "DiaMaJacobi"
                  Exit Sub
               Else
                  tanteta = Sgn(j% - i%) * Wja(i%, j%) / (Wja(i%, i%) - Wja(j%, j%))
               End If
            Else
               discriminant = (Wja(i%, i%) - Wja(j%, j%)) ^ 2 + 4 * Wja(i%, j%) ^ 2
               tanteta = (Sgn(i% - j%) * (Wja(i%, i%) - Wja(j%, j%)) + Sqr(discriminant)) / 2 / Wja(j%, i%)
            End If
            denom = Sqr(1 + tanteta * tanteta)
            sinteta = tanteta / denom
            costeta = 1 / denom
            ' --------------------------------------------------------
            ' construction de Qja et QINVja
            ' --------------------------------------------------------
            For iloc% = 1 To OrdreJa%
               For jloc% = iloc% To OrdreJa%
                  If jloc% = iloc% Then
                     Qja(iloc%, jloc%) = 1
                     QINVja(iloc%, jloc%) = 1
                  Else
                     Qja(iloc%, jloc%) = 0
                     QINVja(iloc%, jloc%) = 0
                     Qja(jloc%, iloc%) = 0
                     QINVja(jloc%, iloc%) = 0
                  End If
               Next jloc%
            Next iloc%
            Qja(i%, i%) = costeta
            QINVja(i%, i%) = costeta
            Qja(i%, j%) = -sinteta
            QINVja(i%, j%) = sinteta
            Qja(j%, i%) = sinteta
            QINVja(j%, i%) = -sinteta
            Qja(j%, j%) = costeta
            QINVja(j%, j%) = costeta
            ' --------------------------------------------------------
            ' calcul de Wja : Wja = QINVja * Wja * Qja
            ' 1ère étape : MKja = Wja * Qja
            ' --------------------------------------------------------
            For iloc% = 1 To OrdreJa%
               For jloc% = 1 To OrdreJa%
                  elem = 0
                  For kloc% = 1 To OrdreJa%
                     elem = elem + Wja(iloc%, kloc%) * Qja(kloc%, jloc%)
                  Next kloc%
                  MKja(iloc%, jloc%) = elem
               Next jloc%
            Next iloc%
            ' --------------------------------------------------------
            ' 2ème étape : Wja = QINVja * MKja
            ' --------------------------------------------------------
            For iloc% = 1 To OrdreJa%
               For jloc% = 1 To OrdreJa%
                  elem = 0
                  For kloc% = 1 To OrdreJa%
                     elem = elem + QINVja(iloc%, kloc%) * MKja(kloc%, jloc%)
                  Next kloc%
                  Wja(iloc%, jloc%) = elem
               Next jloc%
            Next iloc%
            ' --------------------------------------------------------
            ' calcul de Pja : Pja = Pja * Qja
            ' 1ère étape : MKja = Pja * Qja
            ' --------------------------------------------------------
            For iloc% = 1 To OrdreJa%
               For jloc% = 1 To OrdreJa%
                  elem = 0
                  For kloc% = 1 To OrdreJa%
                     elem = elem + Pja(iloc%, kloc%) * Qja(kloc%, jloc%)
                  Next kloc%
                  MKja(iloc%, jloc%) = elem
               Next jloc%
            Next iloc%
            ' --------------------------------------------------------
            ' 2ème étape : Pja = MKja
            ' --------------------------------------------------------
            For iloc% = 1 To OrdreJa%
               For jloc% = 1 To OrdreJa%
                  Pja(iloc%, jloc%) = MKja(iloc%, jloc%)
               Next jloc%
            Next iloc%
            ' --------------------------------------------------------
         End If
      Next j%
   Next i%
   ' -----------------------------------------------------
   ' valeurs propres :
   ' -----------------------------------------------------
   For i% = 1 To OrdreJa%
      Vja(i%) = Wja(i%, i%)
   Next i%
   ' -----------------------------------------------------
   ' facteur de variation des valeurs propres
   ' -----------------------------------------------------
   factvarmax = 0
   If numcycle% = 1 Then
      factvarmax = 10 * eps
   Else
      For i% = 1 To OrdreJa%
         If Vja0(i%) = 0 Then
            factvar = Vja(i%)
         Else
            factvar = Abs(Vja(i%) - Vja0(i%)) / Abs(Vja0(i%))
         End If
         If factvar > factvarmax Then
            factvarmax = factvar
         End If
      Next i%
   End If
   ' -----------------------------------------------------
   ' valeurs propres du cycle précédent :
   ' -----------------------------------------------------
   For i% = 1 To OrdreJa%
      Vja0(i%) = Vja(i%)
   Next i%
   ' -----------------------------------------------------
   ' facteur de couplage maximum
   ' -----------------------------------------------------
   factcouplmax = 0
   For i% = 1 To OrdreJa% - 1
      For j% = i% + 1 To OrdreJa%
         ' facteur de couplage
         If Wja(i%, i%) = 0 Or Wja(j%, j%) = 0 Then
            factcoupl = 10 * Abs(Wja(i%, j%))
         Else
            factcoupl = Abs(Wja(i%, j%)) / Sqr(Abs(Wja(i%, i%) * Wja(j%, j%)))
         End If
         If factcoupl > factcouplmax Then
            factcouplmax = factcoupl
         End If
      Next j%
   Next i%
   ' -----------------------------------------------------
   ' tests de convergence
   ' -----------------------------------------------------
   If factcouplmax < eps And factvarmax < eps Then
      ' Convergence au bout de numcycle% cycles.
      Exit Do
   End If
   ' -----------------------------------------------------
   ' sortie après numcyclesmax% cycles
   ' -----------------------------------------------------
   If numcycle% = numcyclesmax% Then
      Message$ = "Pas de convergence après " & numcycle% & " cycles."
      MsgBox Message$, 16, "DiaMaJacobi"
      Exit Sub
   End If
   ' -----------------------------------------------------
Loop
' -----------------------
'  Valeurs propres :
' -----------------------
For i% = 1 To OrdreJa%
   For j% = 1 To OrdreJa%
      If i% = j% Then
         Wja(i%, j%) = Vja(i%)
      Else
         Wja(i%, j%) = 0
      End If
   Next j%
Next i%
' ---------------------------------------------------
End Sub
Public Sub DiaMaShiftedQR(OrdreQr%, Gqr(), Pmg(), Wqr(), Pqr())
' ***********************************************
'  Recherche des éléments propres d'une matrice
'  réelle par l'algorithme du QR avec déplacement
'  (valeurs propres réelles seulement)
' ***********************************************
' En entrée :
' OrdreQr%    = ordre des matrices en jeu
' Gqr()       = matrice de départ M déjà sous forme de la
'               matrice de HESSENBERG Gqr
' Pmg()       = matrice de passage de M à Gqr :
' Gqr = (transposée de Pmg) x M x Pmg
' En sortie :
' Wqr()       = matrice diagonale
' Pqr()       = matrice unitaire de passage de M à Wqr
'               (vecteurs propres en colonnes)
' -------------------------------------------------------------------
eps = 0.000001
eps2 = 0.0001
' ---------------------------------------------------
' Examen de OrdreQr%
' ---------------------------------------------------
If OrdreQr% < 3 Then
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "DiaMaQR"
End If
' ---------------------------------------------------
' DIMs
' ---------------------------------------------------
'ReDim Gqr(1 To OrdreQr%, 1 To OrdreQr%)  ' matrice de Hessenberg supérieure
ReDim Hqr(1 To OrdreQr%, 1 To OrdreQr%)   ' matrice de Householder unitaire
ReDim MKqr(1 To OrdreQr%, 1 To OrdreQr%)  ' matrice de travail
'ReDim Pmg(1 To OrdreQr%, 1 To OrdreQr%)  ' matrice de passage de M à G
ReDim Pgt(1 To OrdreQr%, 1 To OrdreQr%)   ' matrice de passage de G à (G triangulaire)
ReDim Qqr(1 To OrdreQr%, 1 To OrdreQr%)   ' matrice unitaire
ReDim Rqr(1 To OrdreQr%, 1 To OrdreQr%)   ' matrice triangulaire supérieure
ReDim Vqr(1 To OrdreQr%)                  ' vecteur
ReDim Uqr(1 To OrdreQr%, 1 To OrdreQr%)   ' matrice des vecteurs propres de G
' *****************************************************
' On suppose que M a été mise sous forme
' d'une matrice de HESSENBERG G (par la Sub Hessenberg)
' *****************************************************
' A ce stade :
' G = (transposée de Pmg) x M x Pmg
' *********************************
' Algorithme du QR avec translation
' *********************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
nBoucle% = 100
OrdreSousMat% = OrdreQr% + 1
' ---------------------------------------------------
' Initialisation de Pgt
' ---------------------------------------------------
For iloc% = 1 To OrdreQr%
   For jloc% = 1 To OrdreQr%
      If iloc% = jloc% Then
         Pgt(iloc%, jloc%) = 1
      Else
         Pgt(iloc%, jloc%) = 0
      End If
   Next jloc%
Next iloc%
' ---------------------------------------------------
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' Boucle sur chaque valeur propre
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Do
   OrdreSousMat% = OrdreSousMat% - 1
   iBoucle% = 0
   ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
   ' Boucle principale avec test de convergence
   ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
   Do
      iBoucle% = iBoucle% + 1
      ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      '    Factorisation QR de la sous-matrice de G d'ordre ordreSousMat%
      ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      ' ------------------------------------------------------
      ' Initialisation de Q
      ' ------------------------------------------------------
      For iloc% = 1 To OrdreQr%
         For jloc% = 1 To OrdreQr%
            If iloc% = jloc% Then
               Qqr(iloc%, jloc%) = 1
            Else
               Qqr(iloc%, jloc%) = 0
            End If
         Next jloc%
      Next iloc%
      ' ------------------------------------------------------
      ' Initialisation de R
      ' ------------------------------------------------------
      sk = Gqr(OrdreSousMat%, OrdreSousMat%)
      For iloc% = 1 To OrdreQr%
         For jloc% = 1 To OrdreQr%
            If iloc% = jloc% Then
               Rqr(iloc%, jloc%) = Gqr(iloc%, jloc%) - sk
            Else
               Rqr(iloc%, jloc%) = Gqr(iloc%, jloc%)
            End If
         Next jloc%
      Next iloc%
      ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++
      '           Etape k
      ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++
      For kloc% = 1 To OrdreSousMat% - 1
         ' ---------------------------------------------------
         ' Initialisation de la matrice de Householder Hqr() à I
         ' ---------------------------------------------------
         For iloc% = 1 To OrdreQr%
            For jloc% = 1 To OrdreQr%
               If iloc% = jloc% Then
                  Hqr(iloc%, jloc%) = 1
               Else
                  Hqr(iloc%, jloc%) = 0
               End If
            Next jloc%
         Next iloc%
         ' ----------------------------
         ' Carré et norme du vecteur Ak
         ' ----------------------------
         CarreAk = 0
         NormeAk = 0
         Relkp1 = Rqr(kloc% + 1, kloc%)
         CarreAk = CarreAk + Relkp1 * Relkp1
         If CarreAk < eps Then
            ' H = I
         Else
            Relkk = Rqr(kloc%, kloc%)
            CarreAk = CarreAk + Relkk * Relkk
            NormeAk = Sqr(CarreAk)
            TAUloc = CarreAk + Sgn(Relkk) * Relkk * NormeAk
            ' ---------------------------------------
            ' Calcul du vecteur Vqr()
            ' ---------------------------------------
            Vqr(kloc%) = Relkk + Sgn(Relkk) * NormeAk
            For iloc% = kloc% + 1 To OrdreSousMat%
               Vqr(iloc%) = Rqr(iloc%, kloc%)
            Next iloc%
            ' ---------------------------------------
            ' Calcul de la matrice de Householder Hqr()
            ' ---------------------------------------
            For iloc% = kloc% To OrdreSousMat%
               For jloc% = kloc% To OrdreSousMat%
                  Hqr(iloc%, jloc%) = Hqr(iloc%, jloc%) - Vqr(iloc%) * Vqr(jloc%) / TAUloc
               Next jloc%
            Next iloc%
            ' ---------------------------------------
            ' Calcul de la matrice Rqr()
            ' ---------------------------------------
            For iloc% = 1 To OrdreSousMat%
               For jloc% = 1 To OrdreSousMat%
                  elem = 0
                  For lloc% = 1 To OrdreSousMat%
                     elem = elem + Hqr(iloc%, lloc%) * Rqr(lloc%, jloc%)
                  Next lloc%
                  MKqr(iloc%, jloc%) = elem
               Next jloc%
            Next iloc%
            For iloc% = 1 To OrdreSousMat%
               For jloc% = 1 To OrdreSousMat%
                  Rqr(iloc%, jloc%) = MKqr(iloc%, jloc%)
               Next jloc%
            Next iloc%
            ' ---------------------------------------
            ' Calcul de la matrice Qqr()
            ' ---------------------------------------
            For iloc% = 1 To OrdreSousMat%
               For jloc% = 1 To OrdreSousMat%
                  elem = 0
                  For lloc% = 1 To OrdreSousMat%
                     elem = elem + Qqr(iloc%, lloc%) * Hqr(lloc%, jloc%)
                  Next lloc%
                  MKqr(iloc%, jloc%) = elem
               Next jloc%
            Next iloc%
            For iloc% = 1 To OrdreSousMat%
               For jloc% = 1 To OrdreSousMat%
                  Qqr(iloc%, jloc%) = MKqr(iloc%, jloc%)
               Next jloc%
            Next iloc%
            ' ---------------------------------------
         End If
      Next kloc%
      ' +++++++++++++++++++++++++++++++++++++++++++++
      ' ---------------------------------------
      ' Calcul de la matrice Pgt()
      ' ---------------------------------------
      For iloc% = 1 To OrdreQr%
         For jloc% = 1 To OrdreQr%
            elem = 0
            For lloc% = 1 To OrdreQr%
               elem = elem + Pgt(iloc%, lloc%) * Qqr(lloc%, jloc%)
            Next lloc%
            MKqr(iloc%, jloc%) = elem
         Next jloc%
      Next iloc%
      For iloc% = 1 To OrdreQr%
         For jloc% = 1 To OrdreQr%
            Pgt(iloc%, jloc%) = MKqr(iloc%, jloc%)
         Next jloc%
      Next iloc%
      ' ----------------------------------------------------------
      ' Calcul de la matrice Gk+1() = (transposée de Qk) x Gk x Qk
      ' ----------------------------------------------------------
      For iloc% = 1 To OrdreQr%
         For jloc% = 1 To OrdreQr%
            elem = 0
            For lloc% = 1 To OrdreQr%
               elem = elem + Gqr(iloc%, lloc%) * Qqr(lloc%, jloc%)
            Next lloc%
            MKqr(iloc%, jloc%) = elem
         Next jloc%
      Next iloc%
      For iloc% = 1 To OrdreQr%
         For jloc% = 1 To OrdreQr%
            elem = 0
            For lloc% = 1 To OrdreQr%
               elem = elem + Qqr(lloc%, iloc%) * MKqr(lloc%, jloc%)
            Next lloc%
            Gqr(iloc%, jloc%) = elem
         Next jloc%
      Next iloc%
      ' ----------------------------------------------------------
      ' Test de convergence
      ' ----------------------------------------------------------
      If Abs(Gqr(OrdreSousMat%, OrdreSousMat% - 1)) < eps2 Then
         Exit Do
      End If
      ' ----------------------------------------------------------
      ' Test en cas de non convergence
      ' ----------------------------------------------------------
      If iBoucle% > nBoucle% Then
         Message$ = "Pas de convergence après " & nBoucle% & " boucles"
         MsgBox Message$, 16, "DiaMaShiftedQR"
         Exit Sub
      End If
   Loop
   ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Loop Until OrdreSousMat% = 2
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' +++++++++++++++++++++++++++++++++++++++++++++
' ---------------------------------------
' Calcul de la matrice Qqr() finale
' ---------------------------------------
For iloc% = 1 To OrdreQr%
   For jloc% = 1 To OrdreQr%
      Qqr(iloc%, jloc%) = Pgt(iloc%, jloc%)
   Next jloc%
Next iloc%
' -------------------------------------------------------
' A ce stade :
' G = (transposée de Pgt) x (transposée de Pmg) x M x Pmg x Pgt
' G est triangulaire supérieure d'éléments diagonaux
' les valeurs propres de M
' ---------------------------------------
' Calcul de la matrice Wqr() diagonale
' ---------------------------------------
For iloc% = 1 To OrdreQr%
   For jloc% = 1 To OrdreQr%
      If iloc% = jloc% Then
         Wqr(iloc%, jloc%) = Gqr(iloc%, jloc%)
      Else
         Wqr(iloc%, jloc%) = 0
      End If
   Next jloc%
Next iloc%
' *********************************
' Calcul des vecteurs propres
' *********************************
' ------------------------------------------------------
' Uqr(i,j) = j-ème composante du vecteur propre Ui de G
' correspondant à la valeur propre Gqr(i,i)
' ------------------------------------------------------
Uqr(1, 1) = 1
For jloc% = 2 To OrdreQr%
   Uqr(1, jloc%) = 0
Next jloc%
For iloc% = 2 To OrdreQr%
   For jloc% = OrdreQr% To iloc% + 1 Step -1
      Uqr(iloc%, jloc%) = 0
   Next jloc%
   Uqr(iloc%, iloc%) = 1
   For jloc% = iloc% - 1 To 1 Step -1
      elem = 0
      For kloc% = jloc% + 1 To iloc%
         elem = elem + Gqr(jloc%, kloc%) * Uqr(iloc%, kloc%)
      Next kloc%
      Uqr(iloc%, jloc%) = elem / (Gqr(iloc%, iloc%) - Gqr(jloc%, jloc%))
   Next jloc%
Next iloc%
' ------------------------------------------------------
' Pqr(i,j) = j-ème composante du vecteur propre Vi de M
' correspondant à la valeur propre Gqr(i,i)
' Vi = Pmg x Pgt x Ui
' ------------------------------------------------------
For iloc% = 1 To OrdreQr%
   For jloc% = 1 To OrdreQr%
      elem = 0
      For kloc% = 1 To OrdreQr%
         elem = elem + Pmg(iloc%, kloc%) * Pgt(kloc%, jloc%)
      Next kloc%
      MKqr(iloc%, jloc%) = elem
   Next jloc%
Next iloc%
For iloc% = 1 To OrdreQr%
   For jloc% = 1 To OrdreQr%
      elem = 0
      For kloc% = 1 To OrdreQr%
         elem = elem + MKqr(jloc%, kloc%) * Uqr(iloc%, kloc%)
      Next kloc%
      Pqr(jloc%, iloc%) = elem
   Next jloc%
Next iloc%
' ----------------------------------
' Normalisation des vecteurs propres
' ----------------------------------
For jloc% = 1 To OrdreQr%
   elem = 0
   For iloc% = 1 To OrdreQr%
      elem = elem + Pqr(iloc%, jloc%) * Pqr(iloc%, jloc%)
   Next iloc%
   elem = Sqr(elem)
   For iloc% = 1 To OrdreQr%
      Pqr(iloc%, jloc%) = Pqr(iloc%, jloc%) / elem
   Next iloc%
Next jloc%
' ---------------------------------------------------
End Sub

Public Sub DiaMaQR2VapVep(OrdreDqr%, NbComplexe%, Gdqr(), PMGdqr(), Pdqr(), VaPRdqr(), VaPIdqr())
' *****************************************************
' Recherche des valeurs propres et les vecteurs
' propres réels et complexes d'une matrice réelle
' par l'algorithme du double QR
' *********************************************
' En entrée :
' OrdreDqr%    = ordre des matrices en jeu
' Gdqr()       = matrice de départ M déjà sous forme de la
'                matrice de HESSENBERG Gdqr
' PMGdqr()     = matrice de passage de M à Gdqr :
' Gdqr = (transposée de PMGdqr) x M x PMGdqr

' En sortie :
' NbComplexe%  = nombre de valeurs propres complexes
' VaPRdqr()    = partie réelle des valeurs propres
' VaPIdqr()    = partie imaginaire des valeurs propres
' Pdqr()       = matrice unitaire de passage de M
'                à la matrice des valeurs propres
'               (vecteurs propres en colonnes; si la valeur propre
'               numéro i est complexe, les colonnes i et i+1 de Pdqr()
'               contiennent les parties réelle et imaginaire du
'               vecteur propre correspondant).
' -------------------------------------------------------------------
' Examen de OrdreDqr%
' ---------------------------------------------------
If OrdreDqr% < 3 Then
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "DiaMaQR2VapVep"
End If
' ---------------------------------------------------
' DIMs
' ---------------------------------------------------
ReDim GKdqr(1 To OrdreDqr%, 1 To OrdreDqr%)  ' matrice de Hessenberg supérieure
'                                            ' (matrice de travail)
'ReDim VaPRdqr(1 To OrdreDqr%)                ' partie réelle des valeurs propres
'ReDim VaPIdqr(1 To OrdreDqr%)                ' partie imaginaire des valeurs propres
' *****************************************************
' On suppose que M a été mise sous forme
' d'une matrice de HESSENBERG G (par la Sub Hessenberg)
' *****************************************************
' Algorithme du double QR avec translation
' ****************************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
' On met Gdqr dans la matrice de travail GKdqr
'     et PMGdqr dans la matrice Pdqr
' pour ne pas les modifier en sortie
For i% = 1 To OrdreDqr%
   For j% = 1 To OrdreDqr%
      GKdqr(i%, j%) = Gdqr(i%, j%)
      Pdqr(i%, j%) = PMGdqr(i%, j%)
   Next j%
Next i%
' ---------------------------------------------------
ierr% = 0
norm = 0
k% = 1
NbComplexe% = 0
' ---------------------------------------------------
' Calcul de la norme de la matrice
' ---------------------------------------------------
For i% = 1 To OrdreDqr%
   For j% = k% To OrdreDqr%
      norm = norm + Abs(GKdqr(i%, j%))
   Next j%
   k% = i%
Next i%
' ---------------------------------------------------
en% = OrdreDqr%
itn% = 30 * OrdreDqr%
Tloc = 0
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' Recherche de chaque valeur propre
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Do
   If en% < 1 Then
      Exit Do
   End If
   its% = 0
   na% = en% - 1
   enm2% = na% - 1
   ' ---------------------------------------------------
   ' Recherche un petit élément sous-diagonal isolé
   ' pour lloc% = en% pas -1 jusqu'à 1
   ' ---------------------------------------------------
   Do
      For lloc% = en% To 1 Step -1
         If lloc% = 1 Then
            Exit For
         End If
         Sloc = Abs(GKdqr(lloc% - 1, lloc% - 1)) + Abs(GKdqr(lloc%, lloc%))
         If Sloc = 0 Then
            Sloc = norm
         End If
         tst1 = Sloc
         tst2 = tst1 + Abs(GKdqr(lloc%, lloc% - 1))
         If tst2 = tst1 Then
            Exit For
         End If
      Next lloc%
      ' ---------------------------------------------------
      ' forme un déplacement ("shift")
      ' ---------------------------------------------------
      Xloc = GKdqr(en%, en%)
      If lloc% = en% Then
         Exit Do
      End If
      Yloc = GKdqr(na%, na%)
      Wloc = GKdqr(en%, na%) * GKdqr(na%, en%)
      If lloc% = na% Then
         Exit Do
      End If
      If itn% = 0 Then
         Erreur = True
         Message$ = " Pas de convergence après " & 30 * OrdreDqr% & " itérations" & Chr$(13)
         Message$ = Message$ & " pour la valeur propre numéro " & en%
         MsgBox Message$, 16, "DiaMaQR2VapVep"
         Exit Sub
      End If
      If its% = 10 Or its% = 20 Then
         ' ---------------------------------------------------
         ' forme un déplacement exceptionnel
         ' ---------------------------------------------------
         Tloc = Tloc + Xloc
         For i% = 1 To en%
            GKdqr(i%, i%) = GKdqr(i%, i%) - Xloc
         Next i%
         Sloc = Abs(GKdqr(en%, na%)) + Abs(GKdqr(na%, enm2%))
         Xloc = 0.75 * Sloc
         Yloc = Xloc
         Wloc = -0.4375 * Sloc * Sloc
      End If
      ' ---------------------------------------------------
      its% = its% + 1
      itn% = itn% - 1
      ' ---------------------------------------------------
      ' Recherche deux petits éléments sous-diagonaux
      ' consécutifs pour mloc% = en%-2 pas -1 jusqu'à lloc%
      ' ---------------------------------------------------
      For mloc% = enm2% To lloc% Step -1
         Zloc = GKdqr(mloc%, mloc%)
         Rloc = Xloc - Zloc
         Sloc = Yloc - Zloc
         Ploc = (Rloc * Sloc - Wloc) / GKdqr(mloc% + 1, mloc%) + GKdqr(mloc%, mloc% + 1)
         Qloc = GKdqr(mloc% + 1, mloc% + 1) - Zloc - Rloc - Sloc
         Rloc = GKdqr(mloc% + 2, mloc% + 1)
         Sloc = Abs(Ploc) + Abs(Qloc) + Abs(Rloc)
         Ploc = Ploc / Sloc
         Qloc = Qloc / Sloc
         Rloc = Rloc / Sloc
         If mloc% = lloc% Then
            Exit For
         End If
         tst1 = Abs(Ploc) * (Abs(GKdqr(mloc% - 1, mloc% - 1)) + Abs(Zloc) + Abs(GKdqr(mloc% + 1, mloc% + 1)))
         tst2 = tst1 + Abs(GKdqr(mloc%, mloc% - 1)) * (Abs(Qloc) + Abs(Rloc))
         If tst2 = tst1 Then
            Exit For
         End If
      Next mloc%
      mp2% = mloc% + 2
      For i% = mp2% To en%
         GKdqr(i%, i% - 2) = 0
         If i% <> mp2% Then
            GKdqr(i%, i% - 3) = 0
         End If
      Next i%
      ' ---------------------------------------------------
      ' Etape de double QR impliquant les lignes lloc% à en%
      ' et les colonnes m% à en%
      ' ---------------------------------------------------
      For k% = mloc% To na%
         If k% <> mloc% Then
            Ploc = GKdqr(k%, k% - 1)
            Qloc = GKdqr(k% + 1, k% - 1)
            Rloc = 0
            If k% <> na% Then
               Rloc = GKdqr(k% + 2, k% - 1)
            End If
            Xloc = Abs(Ploc) + Abs(Qloc) + Abs(Rloc)
            If Xloc <> 0 Then
               Ploc = Ploc / Xloc
               Qloc = Qloc / Xloc
               Rloc = Rloc / Xloc
            End If
         End If
         If Xloc <> 0 Then
            Sloc = Sgn(Ploc) * Sqr(Ploc * Ploc + Qloc * Qloc + Rloc * Rloc)
            If k% = mloc% Then
               If lloc% <> mloc% Then
                  GKdqr(k%, k% - 1) = -GKdqr(k%, k% - 1)
               End If
            Else
               GKdqr(k%, k% - 1) = -Sloc * Xloc
            End If
            Ploc = Ploc + Sloc
            Xloc = Ploc / Sloc
            Yloc = Qloc / Sloc
            Zloc = Rloc / Sloc
            Qloc = Qloc / Ploc
            Rloc = Rloc / Ploc
            If k% = na% Then
               ' ---------------------------------------------------
               ' Modification de ligne
               ' ---------------------------------------------------
               For j% = k% To en%
                  Ploc = GKdqr(k%, j%) + Qloc * GKdqr(k% + 1, j%)
                  GKdqr(k%, j%) = GKdqr(k%, j%) - Ploc * Xloc
                  GKdqr(k% + 1, j%) = GKdqr(k% + 1, j%) - Ploc * Yloc
               Next j%
               If en% < k% + 3 Then
                  j% = en%
               Else
                  j% = k% + 3
               End If
               ' ---------------------------------------------------
               ' Modification de colonne
               ' ---------------------------------------------------
               For i% = 1 To j%
                  Ploc = Xloc * GKdqr(i%, k%) + Yloc * GKdqr(i%, k% + 1)
                  GKdqr(i%, k%) = GKdqr(i%, k%) - Ploc
                  GKdqr(i%, k% + 1) = GKdqr(i%, k% + 1) - Ploc * Qloc
               Next i%
               ' ---------------------------------------------------
               ' Accumulation des transformations
               ' ---------------------------------------------------
               For i% = 1 To OrdreDqr%
                  Ploc = Xloc * Pdqr(i%, k%) + Yloc * Pdqr(i%, k% + 1)
                  Pdqr(i%, k%) = Pdqr(i%, k%) - Ploc
                  Pdqr(i%, k% + 1) = Pdqr(i%, k% + 1) - Ploc * Qloc
               Next i%
               ' ---------------------------------------------------
            Else
               ' ---------------------------------------------------
               ' Modification de ligne
               ' ---------------------------------------------------
               For j% = k% To en%
                  Ploc = GKdqr(k%, j%) + Qloc * GKdqr(k% + 1, j%) + Rloc * GKdqr(k% + 2, j%)
                  GKdqr(k%, j%) = GKdqr(k%, j%) - Ploc * Xloc
                  GKdqr(k% + 1, j%) = GKdqr(k% + 1, j%) - Ploc * Yloc
                  GKdqr(k% + 2, j%) = GKdqr(k% + 2, j%) - Ploc * Zloc
               Next j%
               If en% < k% + 3 Then
                  j% = en%
               Else
                  j% = k% + 3
               End If
               ' ---------------------------------------------------
               ' Modification de colonne
               ' ---------------------------------------------------
               For i% = 1 To j%
                  Ploc = Xloc * GKdqr(i%, k%) + Yloc * GKdqr(i%, k% + 1) + Zloc * GKdqr(i%, k% + 2)
                  GKdqr(i%, k%) = GKdqr(i%, k%) - Ploc
                  GKdqr(i%, k% + 1) = GKdqr(i%, k% + 1) - Ploc * Qloc
                  GKdqr(i%, k% + 2) = GKdqr(i%, k% + 2) - Ploc * Rloc
               Next i%
               ' ---------------------------------------------------
               ' Accumulation des transformations
               ' ---------------------------------------------------
               For i% = 1 To OrdreDqr%
                  Ploc = Xloc * Pdqr(i%, k%) + Yloc * Pdqr(i%, k% + 1) + Zloc * Pdqr(i%, k% + 2)
                  Pdqr(i%, k%) = Pdqr(i%, k%) - Ploc
                  Pdqr(i%, k% + 1) = Pdqr(i%, k% + 1) - Ploc * Qloc
                  Pdqr(i%, k% + 2) = Pdqr(i%, k% + 2) - Ploc * Rloc
               Next i%
               ' ---------------------------------------------------
            End If
         End If
      Next k%
   Loop
   If lloc% = en% Then
      ' ---------------------------------------------------
      ' Une racine trouvée
      ' ---------------------------------------------------
      GKdqr(en%, en%) = Xloc + Tloc
      VaPRdqr(en%) = Xloc + Tloc
      VaPIdqr(en%) = 0
      en% = na%
   ElseIf lloc% = na% Then
      ' ---------------------------------------------------
      ' Deux racines trouvées
      ' ---------------------------------------------------
      Ploc = (Yloc - Xloc) / 2
      Qloc = Ploc * Ploc + Wloc
      Zloc = Sqr(Abs(Qloc))
      GKdqr(en%, en%) = Xloc + Tloc
      Xloc = Xloc + Tloc
      GKdqr(na%, na%) = Yloc + Tloc
      If Qloc >= 0 Then
         ' ---------------------------------------------------
         ' Paire réelle
         ' ---------------------------------------------------
         Zloc = Ploc + Sgn(Ploc) * Zloc
         VaPRdqr(na%) = Xloc + Zloc
         VaPRdqr(en%) = Xloc + Zloc
         If Zloc <> 0 Then
            VaPRdqr(en%) = Xloc - Wloc / Zloc
         End If
         VaPIdqr(na%) = 0
         VaPIdqr(en%) = 0
         Xloc = GKdqr(en%, na%)
         Sloc = Abs(Xloc) + Abs(Zloc)
         Ploc = Xloc / Sloc
         Qloc = Zloc / Sloc
         Rloc = Sqr(Ploc * Ploc + Qloc * Qloc)
         Ploc = Ploc / Rloc
         Qloc = Qloc / Rloc
         ' ---------------------------------------------------
         ' Modification de ligne
         ' ---------------------------------------------------
         For j% = na% To OrdreDqr%
            Zloc = GKdqr(na%, j%)
            GKdqr(na%, j%) = Qloc * Zloc + Ploc * GKdqr(en%, j%)
            GKdqr(en%, j%) = Qloc * GKdqr(en%, j%) - Ploc * Zloc
         Next j%
         ' ---------------------------------------------------
         ' Modification de colonne
         ' ---------------------------------------------------
         For i% = 1 To en%
            Zloc = GKdqr(i%, na%)
            GKdqr(i%, na%) = Qloc * Zloc + Ploc * GKdqr(i%, en%)
            GKdqr(i%, en%) = Qloc * GKdqr(i%, en%) - Ploc * Zloc
         Next i%
         ' ---------------------------------------------------
         ' Accumulation des transformations
         ' ---------------------------------------------------
         For i% = 1 To OrdreDqr%
            Zloc = Pdqr(i%, na%)
            Pdqr(i%, na%) = Qloc * Zloc + Ploc * Pdqr(i%, en%)
            Pdqr(i%, en%) = Qloc * Pdqr(i%, en%) - Ploc * Zloc
         Next i%
         ' ---------------------------------------------------
      Else
         ' ---------------------------------------------------
         ' Paire complexe
         ' ---------------------------------------------------
         NbComplexe% = NbComplexe% + 2
         VaPRdqr(na%) = Xloc + Ploc
         VaPRdqr(en%) = Xloc + Ploc
         VaPIdqr(na%) = Zloc
         VaPIdqr(en%) = -Zloc
      End If
         ' ---------------------------------------------------
      en% = enm2%
   End If
Loop
' ---------------------------------------------------
' Toutes les racines ont été trouvées.
' Recherche des vecteurs propres.
' ---------------------------------------------------
If norm = 0 Then
   Exit Sub
End If
' pour en% = OrdreDqr% pas -1 jusqu'à 1
For en% = OrdreDqr% To 1 Step -1
   Ploc = VaPRdqr(en%)
   Qloc = VaPIdqr(en%)
   na% = en% - 1
   If Qloc = 0 Then
      ' -------------------------------------
      ' vecteur réel
      ' -------------------------------------
      mloc% = en%
      GKdqr(en%, en%) = 1
      If na% <> 0 Then
         ' pour i%=  en%-1 pas -1 jusqu'à 1
         For i% = na% To 1 Step -1
            Wloc = GKdqr(i%, i%) - Ploc
            Rloc = 0
            For j% = mloc% To en%
               Rloc = Rloc + GKdqr(i%, j%) * GKdqr(j%, en%)
            Next j%
            If VaPIdqr(i%) < 0 Then
               Zloc = Wloc
               Sloc = Rloc
            Else
               mloc% = i%
               If VaPIdqr(i%) = 0 Then
                  Tloc = Wloc
                  If Tloc = 0 Then
                     tst1 = norm
                     Tloc = tst1
                     Do
                        Tloc = 0.01 * Tloc
                        tst2 = norm + Tloc
                     Loop While tst2 > tst1
                  End If
                  GKdqr(i%, en%) = -Rloc / Tloc
               Else
                  ' résolution des équations réelles
                  Xloc = GKdqr(i%, i% + 1)
                  Yloc = GKdqr(i% + 1, i%)
                  Qloc = (VaPRdqr(i%) - Ploc) ^ 2 + VaPIdqr(i%) ^ 2
                  Tloc = (Xloc * Sloc - Zloc * Rloc) / Qloc
                  GKdqr(i%, en%) = Tloc
                  If Abs(Xloc) > Abs(Zloc) Then
                     GKdqr(i% + 1, en%) = (-Rloc - Wloc * Tloc) / Xloc
                  Else
                     GKdqr(i% + 1, en%) = (-Sloc - Yloc * Tloc) / Zloc
                  End If
               End If
               ' contrôle de dépassement de capacité
               Tloc = Abs(GKdqr(i%, en%))
               If Tloc <> 0 Then
                  tst1 = Tloc
                  tst2 = tst1 + 1 / tst1
                  If tst2 <= tst1 Then
                     For j% = i% To en%
                        GKdqr(j%, en%) = GKdqr(j%, en%) / Tloc
                     Next j%
                  End If
               End If
            End If
         Next i%
         ' fin de vecteur réel
      End If
   ElseIf Qloc < 0 Then
      ' -------------------------------------
      ' vecteur complexe
      ' -------------------------------------
      mloc% = na%
      ' la dernière composante du vecteur est choisie imaginaire
      ' pour que la matrice des vecteurs propres soit triangulaire
      If Abs(GKdqr(en%, na%)) > Abs(GKdqr(na%, en%)) Then
         GKdqr(na%, na%) = Qloc / GKdqr(en%, na%)
         GKdqr(na%, en%) = -(GKdqr(en%, en%) - Ploc) / GKdqr(en%, na%)
      Else
         Call FenetreComplexe.DivisionComplexe(0, -GKdqr(na%, en%), GKdqr(na%, na%) - Ploc, Qloc, GKdqr(na%, na%), GKdqr(na%, en%))
      End If
      GKdqr(en%, na%) = 0
      GKdqr(en%, en%) = 1
      enm2% = na% - 1
      If enm2% <> 0 Then
         ' pour i%=en%-2 pas -1 jusqu'à 1
         For i% = enm2% To 1 Step -1
            Wloc = GKdqr(i%, i%) - Ploc
            RAloc = 0
            SAloc = 0
            For j% = mloc% To en%
               RAloc = RAloc + GKdqr(i%, j%) * GKdqr(j%, na%)
               SAloc = SAloc + GKdqr(i%, j%) * GKdqr(j%, en%)
            Next j%  '760
            If VaPIdqr(i%) < 0 Then
               Zloc = Wloc
               Rloc = RAloc
               Sloc = SAloc
            Else
               mloc% = i%
               If VaPIdqr(i%) = 0 Then
                  Call FenetreComplexe.DivisionComplexe(-RAloc, -SAloc, Wloc, Qloc, GKdqr(i%, na%), GKdqr(i%, en%))
               Else
                  ' résolution des équations complexes
                  Xloc = GKdqr(i%, i% + 1)
                  Yloc = GKdqr(i% + 1, i%)
                  VRloc = (VaPRdqr(i%) - Ploc) ^ 2 + VaPIdqr(i%) ^ 2 - Qloc ^ 2
                  VIloc = (VaPRdqr(i%) - Ploc) * 2 * Qloc
                  If VRloc = 0 And VIloc = 0 Then
                     tst1 = norm * (Abs(Wloc) + Abs(Qloc) + Abs(Xloc) + Abs(Yloc) + Abs(Zloc))
                     VRloc = tst1
                     Do
                        VRloc = 0.01 * VRloc
                        tst2 = tst1 + VRloc
                     Loop While tst2 > tst1
                  End If
                  Call FenetreComplexe.DivisionComplexe(Xloc * Rloc - Zloc * RAloc + Qloc * SAloc, Xloc * Sloc - Zloc * SAloc - Qloc * RAloc, VRloc, VIloc, GKdqr(i%, na%), GKdqr(i%, en%))
                  If Abs(Xloc) > Abs(Zloc) + Abs(Qloc) Then
                     GKdqr(i + 1, na%) = (-RAloc - Wloc * GKdqr(i%, na%) + Qloc * GKdqr(i%, en%)) / Xloc
                     GKdqr(i + 1, en%) = (-SAloc - Wloc * GKdqr(i%, en%) - Qloc * GKdqr(i%, na%)) / Xloc
                  Else
                     Call FenetreComplexe.DivisionComplexe(-Rloc - Yloc * GKdqr(i%, na%), -Sloc - Yloc * GKdqr(i%, en%), Zloc, Qloc, GKdqr(i + 1, na%), GKdqr(i + 1, en%))
                  End If
               End If
               ' contrôle de dépassement de capacité
               Tloc = Abs(GKdqr(i%, na%))
               If Tloc < Abs(GKdqr(i%, en%)) Then
                  Tloc = Abs(GKdqr(i%, en%))
               End If
               If Tloc <> 0 Then
                  tst1 = Tloc
                  tst2 = tst1 + 1 / tst1
                  If tst2 <= tst1 Then
                     For j% = i% To en%
                        GKdqr(j%, na%) = GKdqr(j%, na%) / Tloc
                        GKdqr(j%, en%) = GKdqr(j%, en%) / Tloc
                     Next j%
                  End If
               End If
            End If
         Next i%
         ' fin de vecteur complexe
      End If
   End If
Next en%
' -----------------------------------------------
' Multiplication par la matrice de transformation
' pour donner les vecteurs propres de M
' -----------------------------------------------
' pour j%=OrdreDqr% pas -1 jusqu'à 1
For j% = OrdreDqr% To 1 Step -1
   For i% = 1 To OrdreDqr%
      Zloc = 0
      For k% = 1 To j%
         Zloc = Zloc + Pdqr(i%, k%) * GKdqr(k%, j%)
      Next k%
      Pdqr(i%, j%) = Zloc
   Next i%
Next j%
' --------------------------------------------------
' Normalisation des vecteurs propres
' et remplissage des matrices avec les résultats
' --------------------------------------------------
If NbComplexe% = 0 Then
   ' tous les vecteurs propres réels
   For j% = 1 To OrdreDqr%
      Zloc = 0
      For i% = 1 To OrdreDqr%
         Zloc = Zloc + Pdqr(i%, j%) * Pdqr(i%, j%)
      Next i%
      If Zloc <> 0 Then
         Zloc = Sqr(Zloc)
         For i% = 1 To OrdreDqr%
            Pdqr(i%, j%) = Pdqr(i%, j%) / Zloc
         Next i%
      End If
   Next j%
Else
   ' certains vecteurs propres complexes
   j% = 0
   Do
      j% = j% + 1
      If VaPIdqr(j%) = 0 Then
         Zloc = 0
         For i% = 1 To OrdreDqr%
            Zloc = Zloc + Pdqr(i%, j%) ^ 2
         Next i%
         If Zloc <> 0 Then
            Zloc = Sqr(Zloc)
            For i% = 1 To OrdreDqr%
               Pdqr(i%, j%) = Pdqr(i%, j%) / Zloc
            Next i%
         End If
      Else
         Zloc = 0
         For i% = 1 To OrdreDqr%
            Zloc = Zloc + Pdqr(i%, j%) ^ 2 + Pdqr(i%, j% + 1) ^ 2
         Next i%
         If Zloc <> 0 Then
            Zloc = Sqr(Zloc)
            For i% = 1 To OrdreDqr%
               Pdqr(i%, j%) = Pdqr(i%, j%) / Zloc
               Pdqr(i%, j% + 1) = Pdqr(i%, j% + 1) / Zloc
            Next i%
         End If
         j% = j% + 1
      End If
   Loop Until j% = OrdreDqr%
End If
' -----------------------------------------------
' Libère l'espace mémoire occupé par les tableaux.
' -----------------------------------------------
Erase GKdqr
' -----------------------------------------------
End Sub
Public Sub Hessenberg(OrdreHes%, Mhes(), Ghes(), PermHes%())
' *******************************************************
' Réduction d'une matrice carrée réelle quelconque M
' en matrice de HESSENBERG supérieure G
' telle que : G(i,j) = 0 si i > j+1
' *******************************************************
' En entrée :
' OrdreHes%    = ordre des matrices en jeu
' Mhes()       = matrice de départ

' En sortie :
' Ghes()       = matrice de HESSENBERG; les multiplicateurs utilisés
'                lors de la réduction de M sont placés dans les
'                éléments G(i,j) tels que i > j+1
' PermHes%()   = vecteur de permutation des lignes et colonnes
'                lors du passage de Mhes à Ghes
'                (pouvant servir au calcul ultérieur de la matrice
'                 de passage de M à G)
' -------------------------------------------------------------------
' Examen de OrdreHes%
' ---------------------------------------------------
If OrdreHes% < 3 Then
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "Hessenberg"
End If
' ---------------------------------------------------
' DIMs
' ---------------------------------------------------
'ReDim Ghes(1 To OrdreHes%, 1 To OrdreHes%)  ' matrice de Hessenberg supérieure
'Redim PermHes%(1 To OrdreHes%)
' *****************************************************
'
' *****************************************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
' On met Mhes dans la matrice Ghes
' --------------------------------
For iloc% = 1 To OrdreHes%
   For jloc% = 1 To OrdreHes%
      Ghes(iloc%, jloc%) = Mhes(iloc%, jloc%)
   Next jloc%
Next iloc%
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
la% = OrdreHes% - 1
kp1% = 2
If la% < kp1% Then
   Exit Sub
End If
' ---------------------------------------------------
' Boucle principale
' ---------------------------------------------------
For mloc% = kp1% To la%
   mm1% = mloc% - 1
   Xloc = 0
   iloc% = mloc%
   For jloc% = mloc% To OrdreHes%
      If Abs(Ghes(jloc%, mm1%)) > Abs(Xloc) Then
         Xloc = Ghes(jloc%, mm1%)
         iloc% = jloc%
      End If
   Next jloc%
   PermHes%(mloc%) = iloc%
   If iloc% <> mloc% Then
      ' ---------------------------------------
      ' permutation des lignes et colonnes de M
      ' ---------------------------------------
      For jloc% = mm1% To OrdreHes%
         Yloc = Ghes(iloc%, jloc%)
         Ghes(iloc%, jloc%) = Ghes(mloc%, jloc%)
         Ghes(mloc%, jloc%) = Yloc
      Next jloc%
      For jloc% = 1 To OrdreHes%
         Yloc = Ghes(jloc%, iloc%)
         Ghes(jloc%, iloc%) = Ghes(jloc%, mloc%)
         Ghes(jloc%, mloc%) = Yloc
      Next jloc%
      ' ---------------------------------------
   End If
   If Xloc <> 0 Then
      mp1% = mloc% + 1
      For iloc% = mp1% To OrdreHes%
         Yloc = Ghes(iloc%, mm1%)
         If Yloc <> 0 Then
            Yloc = Yloc / Xloc
            Ghes(iloc%, mm1%) = Yloc
            For jloc% = mloc% To OrdreHes%
               Ghes(iloc%, jloc%) = Ghes(iloc%, jloc%) - Yloc * Ghes(mloc%, jloc%)
            Next jloc%
            For jloc% = 1 To OrdreHes%
               Ghes(jloc%, mloc%) = Ghes(jloc%, mloc%) + Yloc * Ghes(jloc%, iloc%)
            Next jloc%
         End If
      Next iloc% '160
   End If
Next mloc%  '180
' ---------------------------------------------------
End Sub
Public Sub PassMatHes(OrdrePmh%, PMGpmh(), Gpmh(), PermPmh%())
' **************************************************************************
' Calcul de la matrice de passage d'une matrice carrée réelle quelconque M
' à une matrice de HESSENBERG supérieure G telle que : G(i,j) = 0 si i > j+1
' **************************************************************************
' En entrée :
' OrdrePmh%    = ordre des matrices en jeu
' Gpmh()       = matrice de HESSENBERG contenant également les
'                multiplicateurs utilisés pour sa réduction dans ses
'                éléments G(i,j) tels que i > j+1
' PermPmh%()   = vecteur de permutation des lignes et colonnes
'                lors du passage de M à G


' En sortie :
' PMGphm()     = matrice de passage de M à G
' -------------------------------------------------------------------
' Examen de OrdrePmh%
' ---------------------------------------------------
If OrdrePmh% < 3 Then
   Message$ = " L'ordre de la matrice doit être supérieur à 2 !!!"
   MsgBox Message$, 16, "PassMatHes"
End If
' ---------------------------------------------------
' DIMs
' ---------------------------------------------------
'ReDim Gpmh(1 To OrdrePmh%, 1 To OrdrePmh%)  ' matrice de Hessenberg supérieure
'Redim PermPmh%(1 To OrdrePmh%)
' *****************************************************
'
' *****************************************************
' ---------------------------------------------------
' Initialisation de PMGpmh() à la matrice identité I
' ---------------------------------------------------
For iloc% = 1 To OrdrePmh%
   For jloc% = 1 To OrdrePmh%
      PMGpmh(iloc%, jloc%) = 0
   Next jloc%
   PMGpmh(iloc%, iloc%) = 1
Next iloc%
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
kl% = OrdrePmh% - 2
If kl% < 1 Then
   Exit Sub
End If
' ---------------------------------------------------
' Boucle principale
' ---------------------------------------------------
For mp% = OrdrePmh% - 1 To 2 Step -1
   mp1% = mp% + 1
   For iloc% = mp1% To OrdrePmh%
      PMGpmh(iloc%, mp%) = Gpmh(iloc%, mp% - 1)
   Next iloc%
   iloc% = PermPmh%(mp%)
   If iloc% <> mp% Then
      For jloc% = mp% To OrdrePmh%
         PMGpmh(mp%, jloc%) = PMGpmh(iloc%, jloc%)
         PMGpmh(iloc%, jloc%) = 0
      Next jloc%
      PMGpmh(iloc%, mp%) = 1
   End If
Next mp%
' ---------------------------------------------------
End Sub

Private Sub ProduitWxM2xM()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase M2mat, Pmat
   ReDim M1mat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim M2mat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Pmat(1 To OrdreMat%, 1 To OrdreMat%)
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   On Error Resume Next
   DoEvents
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de M2mat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridM2mat.Row = i%
         FenetreMatrice.GridM2mat.Col = j%
         If FenetreMatrice.GridM2mat.Text = "" Then
            FenetreMatrice.GridM2mat.Text = 0
            M2mat(i%, j%) = 0
         Else
            M2mat(i%, j%) = CSng(FenetreMatrice.GridM2mat.Text)
         End If
      Next j%
   Next i%
   On Error GoTo 0
   ' *************************************************
   ' calcul de M1 produit de WxM2
   ' *************************************************
   If GridWmat.Cols = 2 Then
      Message$ = "Produit impossible !"
      MsgBox Message$, 48, "ProduitWxM2xM"
      Exit Sub
   End If
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         M1mat(i%, j%) = 0
         For k% = 1 To OrdreMat%
            M1mat(i%, j%) = M1mat(i%, j%) + Wmat(i%, k%) * M2mat(k%, j%)
         Next k%
      Next j%
   Next i%
   ' *************************************************
   ' calcul de P produit de M1xM
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         Pmat(i%, j%) = 0
         For k% = 1 To OrdreMat%
            Pmat(i%, j%) = Pmat(i%, j%) + M1mat(i%, k%) * Mmat(k%, j%)
         Next k%
      Next j%
   Next i%
   ' *************************************************
   ' affichage du produit effectué
   ' *************************************************
   lblPmat.Caption = "Matrice P produit de WxM2xM :"
   ' *************************************************
   ' affichage des éléments de Pmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         FenetreMatrice.GridPmat.Row = i%
         FenetreMatrice.GridPmat.Col = j%
         FenetreMatrice.GridPmat.Text = Format(Pmat(i%, j%), "0.000")
      Next j%
   Next i%
   ' ********************************************
   lblInfo.Caption = ""
   ' ********************************************
   '***************************************
   '************ Explications *************
   '***************************************
   Info$ = "Dernier calcul effectué : Produit WxM2xM"
   lblInfo.ForeColor = BLEU
   lblInfo.Font.bold = True
   lblInfo.Font.underline = False
   lblInfo.Caption = Info$
   '***************************************
End Sub
