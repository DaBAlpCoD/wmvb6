VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreSysLin 
   Caption         =   "Système linéaire"
   ClientHeight    =   5880
   ClientLeft      =   210
   ClientTop       =   990
   ClientWidth     =   9045
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
   ScaleHeight     =   5880
   ScaleWidth      =   9045
   Begin MSFlexGridLib.MSFlexGrid GridVvec 
      Height          =   2055
      Left            =   7080
      TabIndex        =   11
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3625
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
   Begin MSFlexGridLib.MSFlexGrid GridUvec 
      Height          =   2055
      Left            =   4800
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3625
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
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
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
   Begin VB.PictureBox pctExplic 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.TextBox txtNbEq 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblegale 
      Alignment       =   2  'Center
      Caption         =   "="
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
      TabIndex        =   1
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Caption         =   "x"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   4440
      Width           =   255
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
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblXvec 
      Caption         =   "Vecteur X :"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblYvec 
      Caption         =   "Vecteur Y :"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblNbEq 
      Caption         =   "Nombre d'équations :"
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
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblMmat 
      Caption         =   "Matrice M :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
         Begin VB.Menu mnuOuvrirM 
            Caption         =   "Matrice &M"
         End
         Begin VB.Menu mnuOuvrirY 
            Caption         =   "Vecteur &Y"
         End
      End
      Begin VB.Menu mnuEnregistrer 
         Caption         =   "&Enregistrer..."
         Begin VB.Menu mnuEnregM 
            Caption         =   "Matrice M"
         End
         Begin VB.Menu mnuEnregX 
            Caption         =   "Vecteur X"
         End
         Begin VB.Menu mnuEnregY 
            Caption         =   "Vecteur Y"
         End
      End
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreSysLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NomMatVec$










Private Sub Form_Load()
   '***************************************
   '************ Explications *************
   '***************************************
   ' Mmat(i,j) . Uvec(j) = Vvec(i)
   pctExplic.Cls
   pctExplic.ForeColor = BLEU
   pctExplic.Font.underline = False
   pctExplic.Print " Système de n équations linéaires à n inconnues :"
   pctExplic.Print
   pctExplic.Print " M(1,1)*X(1)+...+M(1,j)*X(j)+...+M(1,n)*X(n) = Y(1)"
   pctExplic.Print " M(2,1)*X(1)+...+M(2,j)*X(j)+...+M(2,n)*X(n) = Y(2)"
   pctExplic.Print " . . . . . . . . ."
   pctExplic.Print " M(i,1)*X(1)+...+M(i,j)*X(j)+...+M(i,n)*X(n) = Y(i)"
   pctExplic.Print " . . . . . . . . ."
   pctExplic.Print " M(n,1)*X(1)+...+M(n,j)*X(j)+...+M(n,n)*X(n) = Y(n)"
   pctExplic.Print
   pctExplic.Print
   pctExplic.Print " Ceci peut aussi s'écrire :  M x X = Y"
   pctExplic.Print " où M est une matrice carrée nxn,"
   pctExplic.Print " et X et Y deux vecteurs à n composantes."
   '***************************************
   OrdreMat% = 3
   txtNbEq.Text = OrdreMat%
   '***************************************
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Uvec(1 To OrdreMat%)
   ReDim Vvec(1 To OrdreMat%)
   '***************************************
   '***** mise en place des 2 grilles *****
   '***************************************
   GridMmat.Rows = OrdreMat% + 1
   GridMmat.Cols = OrdreMat% + 1
   GridUvec.Rows = OrdreMat% + 1
   GridVvec.Rows = OrdreMat% + 1
   '***************************************
   GridMmat.FixedAlignment(0) = 2
   GridMmat.ColWidth(0) = 500
   For i% = 1 To OrdreMat%
      GridMmat.FixedAlignment(i%) = 2
      GridMmat.ColWidth(i%) = 1000
   Next i%
   GridUvec.FixedAlignment(0) = 2
   GridUvec.FixedAlignment(1) = 2
   GridUvec.ColWidth(0) = 500
   GridUvec.ColWidth(1) = 1000
   GridVvec.FixedAlignment(0) = 2
   GridVvec.FixedAlignment(1) = 2
   GridVvec.ColWidth(0) = 500
   GridVvec.ColWidth(1) = 1000
   ' ********************************
   ' numérotation 1ères lignes
   ' ********************************
   GridMmat.Row = 0
   For i% = 1 To OrdreMat%
      GridMmat.Col = i%
      GridMmat.Text = i%
   Next i%
   GridUvec.Row = 0
   GridUvec.Col = 1
   GridUvec.Text = ""
   GridVvec.Row = 0
   GridVvec.Col = 1
   GridVvec.Text = ""
   ' *********************************
   ' numérotation 1ères colonnes
   ' *********************************
   GridMmat.Col = 0
   GridUvec.Col = 0
   GridVvec.Col = 0
   For i% = 1 To OrdreMat%
      GridMmat.Row = i%
      GridMmat.Text = i%
      GridUvec.Row = i%
      GridUvec.Text = i%
      GridVvec.Row = i%
      GridVvec.Text = i%
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
   ' -------------
   Uvec(1) = 0
   Uvec(2) = 0
   Uvec(3) = 0
   ' -------------
   Vvec(1) = 7
   Vvec(2) = 11
   Vvec(3) = 10
   ' -------------
   ' Solution : 1,2,3
   ' ****************************************
   ' placement des valeurs par défaut de Mmat
   ' ****************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         GridMmat.Row = i%
         GridMmat.Col = j%
         GridMmat.Text = Format(Mmat(i%, j%), "0.000")
      Next j%
   Next i%
   ' ****************************************
   ' placement des valeurs par défaut de Uvec
   ' ****************************************
   GridUvec.Col = 1
   For i% = 1 To OrdreMat%
      GridUvec.Row = i%
      GridUvec.Text = ""
   Next i%
   ' ****************************************
   ' placement des valeurs par défaut de Vvec
   ' ****************************************
   GridVvec.Col = 1
   For i% = 1 To OrdreMat%
      GridVvec.Row = i%
      GridVvec.Text = Format(Vvec(i%), "0.000")
   Next i%
   ' ****************************************
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


Private Sub gridVvec_KeyPress(KeyAscii As Integer)
   EleMText$ = GridVvec.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      GridVvec.Text = EleMText$
   Case 8
      If Len(GridVvec.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         GridVvec.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub






Private Sub mnuCalculer_Click()
   Call SolSys
End Sub







Private Sub mnuEnregM_Click()
   NomMatVec$ = "M"
   Call EnregMat
End Sub









Private Sub mnuEnregX_Click()
   NomMatVec$ = "X"
   Call EnregMat
End Sub

Private Sub mnuEnregY_Click()
   NomMatVec$ = "Y"
   Call EnregMat
End Sub


Private Sub mnuOuvrirM_Click()
   NomMatVec$ = "M"
   Call OuvreMat
End Sub





Private Sub mnuOuvrirY_Click()
   NomMatVec$ = "Y"
   Call OuvreMat
End Sub

Private Sub mnuQuitter_Click()
   FenetreSysLin.Hide
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
      Erase Mmat
      ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
      ReDim Vvec(1 To OrdreMat%)
   ElseIf NomMatVec$ = "M" Then
      Erase Mmat
      ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ElseIf NomMatVec$ = "Y" Then
      Erase Vvec
      ReDim Vvec(1 To OrdreMat%)
   End If
   Erase Uvec
   ReDim Uvec(1 To OrdreMat%)
   ' *************************************************
   ' lecture des éléments de la matrice
   ' *************************************************
   For i% = 1 To OrdreMat%
      If NomMatVec$ = "Y" Then
      Input #1, Vvec(i%)
      End If
      For j% = 1 To OrdreMat%
         If NomMatVec$ = "M" Then
            Input #1, Mmat(i%, j%)
         End If
      Next j%
   Next i%
   Close #1
   ' *************************************************
   ' placement de l'ordre des matrices, ce qui provoque
   ' le redimentionnement des grilles
   ' *************************************************
   txtNbEq.Text = Format(OrdreMat%, "0")
   ' *********************************************
   ' placement des nouvelles valeurs de la matrice
   ' *********************************************
   GridUvec.Col = 1
   For i% = 1 To OrdreMat%
      If NomMatVec$ = "Y" Then
         GridUvec.Row = i%
         GridUvec.Text = Format(Uvec(i%), "0.000")
      End If
      For j% = 1 To OrdreMat%
         If NomMatVec$ = "M" Then
            GridMmat.Row = i%
            GridMmat.Col = j%
            GridMmat.Text = Format(Mmat(i%, j%), "0.000")
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
   Maths.ctrlCMDialog.DefaultExt = "mat"
   Maths.ctrlCMDialog.Filter = "Matrice (*.mat)|*.mat"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Mmat, Uvec, Vvec
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Uvec(1 To OrdreMat%)
   ReDim Vvec(1 To OrdreMat%)
   ' *****************************************
   ' affectation de leurs valeurs aux éléments
   ' des matrices et vecteurs
   ' *****************************************
   For i% = 1 To OrdreMat%
      ' -------------------------------
      GridUvec.Row = i%
      GridUvec.Col = 1
      If GridUvec.Text = "" Then
         GridUvec.Text = 0
         Uvec(i%) = 0
      Else
         Uvec(i%) = CSng(GridUvec.Text)
      End If
      ' -------------------------------
      GridVvec.Row = i%
      GridVvec.Col = 1
      If GridVvec.Text = "" Then
         GridVvec.Text = 0
         Vvec(i%) = 0
      Else
         Vvec(i%) = CSng(GridVvec.Text)
      End If
      ' -------------------------------
      For j% = 1 To OrdreMat%
         GridMmat.Row = i%
         GridMmat.Col = j%
         If GridMmat.Text = "" Then
            GridMmat.Text = 0
            Mmat(i%, j%) = 0
         Else
            Mmat(i%, j%) = CSng(GridMmat.Text)
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
            If NomMatVec$ = "X" Then
               Write #1, Uvec(i%)
            ElseIf NomMatVec$ = "Y" Then
               Write #1, Vvec(i%)
            End If
         For j% = 1 To OrdreMat%
            If NomMatVec$ = "M" Then
               Write #1, Mmat(i%, j%)
            End If
         Next j%
      Next i%
   Close #1
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





Private Sub txtNbEq_Change()
   If txtNbEq.Text = "" Then
      OrdreMat% = 1
   Else
      OrdreMat% = CInt(txtNbEq.Text)
   End If
   If OrdreMat% < 1 Then
      Beep
      MsgBox "le nombre d'équations doit être supérieur ou égal à 1 !", 48, "SYSLIN"
      OrdreMat% = 1
      txtNbEq.Text = "1"
   ElseIf OrdreMat% > 30000 Then
      Beep
      MsgBox "le nombre d'équations doit être inférieur à 30000 !", 48, "SYSLIN"
      OrdreMat% = 1
   End If
   GridMmat.Rows = OrdreMat% + 1
   GridMmat.Cols = OrdreMat% + 1
   GridUvec.Rows = OrdreMat% + 1
   GridVvec.Rows = OrdreMat% + 1
   '***************************************
   GridUvec.FixedAlignment(0) = 2
   GridUvec.FixedAlignment(1) = 2
   GridUvec.ColWidth(1) = 1000
   GridVvec.FixedAlignment(0) = 2
   GridVvec.FixedAlignment(1) = 2
   GridVvec.ColWidth(1) = 1000
   GridMmat.FixedAlignment(0) = 2
   For i% = 1 To OrdreMat%
      GridMmat.FixedAlignment(i%) = 2
      GridMmat.ColWidth(i%) = 1000
   Next i%
   ' ********************************
   ' renumérotation 1ère ligne Mmat
   GridMmat.Row = 0
   For i% = 1 To OrdreMat%
      GridMmat.Col = i%
      GridMmat.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère colonne Mmat
   GridMmat.Col = 0
   For i% = 1 To OrdreMat%
      GridMmat.Row = i%
      GridMmat.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère ligne Uvec
   GridUvec.Row = 0
   GridUvec.Col = 1
   GridUvec.Text = ""
   ' ********************************
   ' renumérotation 1ère colonne Uvec
   GridUvec.Col = 0
   For i% = 1 To OrdreMat%
      GridUvec.Row = i%
      GridUvec.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère ligne Vvec
   GridVvec.Row = 0
   GridVvec.Col = 1
   GridVvec.Text = ""
   ' ********************************
   ' renumérotation 1ère colonne Vvec
   GridVvec.Col = 0
   For i% = 1 To OrdreMat%
      GridVvec.Row = i%
      GridVvec.Text = i%
   Next i%
   ' ********************************
End Sub



Public Sub SolSys()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Mmat, Uvec, Vvec
   ReDim Mmat(1 To OrdreMat%, 1 To OrdreMat%)
   ReDim Uvec(1 To OrdreMat%)
   ReDim Vvec(1 To OrdreMat%)
   ' *************************************************
   '               Initialisations
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' *************************************************
   ' remise en place de la grille de Uvec et Vvec
   ' **************************************
   GridUvec.Rows = OrdreMat% + 1
   GridVvec.Rows = OrdreMat% + 1
   '***************************************
   GridUvec.FixedAlignment(0) = 2
   GridUvec.FixedAlignment(1) = 2
   GridUvec.ColWidth(1) = 1000
   GridVvec.FixedAlignment(0) = 2
   GridVvec.FixedAlignment(1) = 2
   GridVvec.ColWidth(1) = 1000
   ' ********************************
   ' renumérotation 1ère ligne Uvec
   GridUvec.Row = 0
   GridUvec.Col = 1
   GridUvec.Text = ""
   ' ********************************
   ' renumérotation 1ère colonne Uvec
   GridUvec.Col = 0
   For i% = 1 To OrdreMat%
      GridUvec.Row = i%
      GridUvec.Text = i%
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Mmat
   ' *************************************************
   For i% = 1 To OrdreMat%
      For j% = 1 To OrdreMat%
         GridMmat.Row = i%
         GridMmat.Col = j%
         If GridMmat.Text = "" Then
            GridMmat.Text = 0
            Mmat(i%, j%) = 0
         Else
            Mmat(i%, j%) = CSng(GridMmat.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Vvec
   ' *************************************************
   GridVvec.Col = 1
   For i% = 1 To OrdreMat%
      GridVvec.Row = i%
      If GridVvec.Text = "" Then
         GridVvec.Text = 0
         Vvec(i%) = 0
      Else
         Vvec(i%) = CSng(GridVvec.Text)
      End If
   Next i%
   ' *************************************************
   ' calcul de Uvec par Triangulation
   ' *************************************************
   Call TrianMat
   If Erreur = True Then
      Message$ = "Système insoluble"
      MsgBox Message$, 48
      Exit Sub
   End If
   ' *************************************************
   ' affichage des éléments de Uvec
   ' *************************************************
   GridUvec.Col = 1
   For i% = 1 To OrdreMat%
      GridUvec.Row = i%
      GridUvec.Text = Format(Uvec(i%), "0.000")
   Next i%
   ' *************************************************
   lblInfo.Caption = ""
   ' *************************************************
End Sub
