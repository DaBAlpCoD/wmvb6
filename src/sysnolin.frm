VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreSysNonLin 
   Caption         =   "Système non linéaire"
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
   Begin MSFlexGridLib.MSFlexGrid GridUpar 
      Height          =   2775
      Left            =   6840
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4895
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
   Begin MSFlexGridLib.MSFlexGrid GridSENL 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4895
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
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   5715
      TabIndex        =   7
      Top             =   120
      Width           =   5775
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
      TabIndex        =   1
      Top             =   2280
      Width           =   735
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
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblInc 
      Caption         =   "Inconnues :"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   2640
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblMmat 
      Caption         =   "Système :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreSysNonLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoSNLMatVec$









Private Sub Form_Load()
   '***************************************
   '************ Explications *************
   '***************************************
   pctExplic.Cls
   pctExplic.ForeColor = BLEU
   pctExplic.Font.underline = False
   pctExplic.Print " Système de n équations non linéaires à n inconnues :"
   pctExplic.Print
   pctExplic.Print " F1[U(1),U(2),...,U(n)] = 0"
   pctExplic.Print " F2[U(1),U(2),...,U(n)] = 0"
   pctExplic.Print " . . . . . . . . . . ."
   pctExplic.Print " Fn[U(1),U(2),...,U(n)] = 0"
   pctExplic.Print
   pctExplic.Print " Où U(1),U(2),...,U(n) sont les inconnues"
   pctExplic.Print " et F1,F2,...,Fn des fonctions de ces inconnues."
   '***************************************
   NbEq% = 3
   txtNbEq.Text = NbEq%
   '***************************************
   ReDim SENL$(1 To NbEq%)
   ReDim Upar(1 To NbEq%)
   '***************************************
   '***** mise en place des 3 grilles *****
   '***************************************
   gridSENL.Rows = NbEq% + 1
   gridUpar.Rows = NbEq% + 1
   '***************************************
   gridSENL.FixedAlignment(0) = 2
   gridSENL.FixedAlignment(1) = 2
   gridSENL.ColWidth(1) = 8000
   gridUpar.FixedAlignment(0) = 2
   gridUpar.FixedAlignment(1) = 2
   gridUpar.ColWidth(1) = 1000
   ' ********************************
   ' numérotation 1ères lignes
   ' ********************************
   gridSENL.Row = 0
   gridSENL.Col = 1
   gridSENL.Text = ""
   gridUpar.Row = 0
   gridUpar.Col = 1
   gridUpar.Text = ""
   ' *********************************
   ' numérotation 1ères colonnes
   ' *********************************
   gridSENL.Col = 0
   gridUpar.Col = 0
   For i% = 1 To NbEq%
      gridSENL.Row = i%
      gridSENL.Text = "F" & i% & " ="
      gridUpar.Row = i%
      gridUpar.Text = "U(" & i% & ") ="
   Next i%
   ' ****************************************
   '       Valeurs par défaut
   ' ****************************************
   SENL$(1) = "6*U(1)+U(2)*U(3)+U(3)-15"
   SENL$(2) = "U(1)^2+7*U(2)-U(3)-12"
   SENL$(3) = "U(1)+U(2)^2-9*U(3)+22"
   ' Solution : U1=1; U2=2; U3=3
   ' -------------
   Upar(1) = 0
   Upar(2) = 0
   Upar(3) = 0
   ' -------------
   ' *****************************************
   ' placement des valeurs par défaut de SENL$
   ' *****************************************
   gridSENL.Col = 1
   For i% = 1 To NbEq%
      gridSENL.Row = i%
      gridSENL.Text = SENL$(i%)
   Next i%
   ' ****************************************
End Sub


Private Sub gridSENL_KeyPress(KeyAscii As Integer)
   EleMText$ = gridSENL.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      gridSENL.Text = EleMText$
   Case 8
      If Len(gridSENL.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         gridSENL.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub



Private Sub mnuCalculer_Click()
   Call SolSysNonLin
End Sub








Private Sub mnuEnregM_Click()
   NoSNLMatVec$ = "M"
   Call EnregMat
End Sub









Private Sub mnuEnregX_Click()
   NoSNLMatVec$ = "X"
   Call EnregMat
End Sub

Private Sub mnuEnregY_Click()
   NoSNLMatVec$ = "Y"
   Call EnregMat
End Sub


Private Sub mnuOuvrirM_Click()
   NoSNLMatVec$ = "M"
   Call OuvreMat
End Sub





Private Sub mnuOuvrirY_Click()
   NoSNLMatVec$ = "Y"
   Call OuvreMat
End Sub

Private Sub mnuQuitter_Click()
   FenetreSysNonLin.Hide
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
   ' et écriture de ces éléments dans SENL(i%,j%)
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, OrdreMatLoc%
   ' ********************************************
   '               ReDims
   ' ********************************************
   If OrdreMatLoc% <> NbEq% Then
      NbEq% = OrdreMatLoc%
      Erase SENL
      ReDim SENL(1 To NbEq%, 1 To NbEq%)
      ReDim Vvec(1 To NbEq%)
   ElseIf NoSNLMatVec$ = "M" Then
      Erase SENL
      ReDim SENL(1 To NbEq%, 1 To NbEq%)
   ElseIf NoSNLMatVec$ = "Y" Then
      Erase Vvec
      ReDim Vvec(1 To NbEq%)
   End If
   Erase Upar
   ReDim Upar(1 To NbEq%)
   ' *************************************************
   ' lecture des éléments de la matrice
   ' *************************************************
   For i% = 1 To NbEq%
      If NoSNLMatVec$ = "Y" Then
      Input #1, Vvec(i%)
      End If
      For j% = 1 To NbEq%
         If NoSNLMatVec$ = "M" Then
            Input #1, SENL(i%, j%)
         End If
      Next j%
   Next i%
   Close #1
   ' *************************************************
   ' placement de l'ordre des matrices, ce qui provoque
   ' le redimentionnement des grilles
   ' *************************************************
   txtNbEq.Text = Format(NbEq%, "0")
   ' *********************************************
   ' placement des nouvelles valeurs de la matrice
   ' *********************************************
   gridUpar.Col = 1
   For i% = 1 To NbEq%
      If NoSNLMatVec$ = "Y" Then
         gridUpar.Row = i%
         gridUpar.Text = Format(Upar(i%), "0.000")
      ElseIf NoSNLMatVec$ = "M" Then
         For j% = 1 To NbEq%
            gridSENL.Row = i%
            gridSENL.Col = j%
            gridSENL.Text = Format(SENL$(i%, j%), "0.000")
         Next j%
      End If
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
   Erase SENL, Upar, Vvec
   ReDim SENL(1 To NbEq%, 1 To NbEq%)
   ReDim Upar(1 To NbEq%)
   ReDim Vvec(1 To NbEq%)
   ' *****************************************
   ' affectation de leurs valeurs aux éléments
   ' des matrices et vecteurs
   ' *****************************************
   For i% = 1 To NbEq%
      ' -------------------------------
      gridUpar.Row = i%
      gridUpar.Col = 1
      If gridUpar.Text = "" Then
         gridUpar.Text = 0
         Upar(i%) = 0
      Else
         Upar(i%) = CSng(gridUpar.Text)
      End If
      ' -------------------------------
      gridVvec.Row = i%
      gridVvec.Col = 1
      If gridVvec.Text = "" Then
         gridVvec.Text = 0
         Vvec(i%) = 0
      Else
         Vvec(i%) = CSng(gridVvec.Text)
      End If
      ' -------------------------------
      For j% = 1 To NbEq%
         gridSENL.Row = i%
         gridSENL.Col = j%
         If gridSENL.Text = "" Then
            gridSENL.Text = 0
            SENL(i%, j%) = 0
         Else
            SENL(i%, j%) = CSng(gridSENL.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   '-----------------------------------------------------
   ' Création du fichier d'éléments de matrice
   ' et écriture de ces éléments dans le fichier
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
      Write #1, NbEq%
      For i% = 1 To NbEq%
         If NoSNLMatVec$ = "X" Then
            Write #1, Upar(i%)
         ElseIf NoSNLMatVec$ = "Y" Then
            Write #1, Vvec(i%)
         ElseIf NoSNLMatVec$ = "M" Then
            For j% = 1 To NbEq%
               Write #1, SENL$(i%, j%)
            Next j%
         End If
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
      NbEq% = 1
   Else
      NbEq% = CInt(txtNbEq.Text)
   End If
   If NbEq% < 1 Then
      Beep
      MsgBox "le nombre d'équations doit être supérieur ou égal à 1 !", 48, "SYSNONLIN"
      NbEq% = 1
      txtNbEq.Text = "1"
   ElseIf NbEq% > 30000 Then
      Beep
      MsgBox "le nombre d'équations doit être inférieur à 30000 !", 48, "SYSNONLIN"
      NbEq% = 1
   End If
   gridSENL.Rows = NbEq% + 1
   gridUpar.Rows = NbEq% + 1
   '***************************************
   gridUpar.FixedAlignment(0) = 2
   gridUpar.FixedAlignment(1) = 2
   gridUpar.ColWidth(1) = 1000
   gridSENL.FixedAlignment(0) = 2
   gridSENL.FixedAlignment(1) = 2
   gridSENL.ColWidth(1) = 8000
   ' ********************************
   ' renumérotation 1ère ligne SENL
   gridSENL.Row = 0
   gridSENL.Col = 1
   gridSENL.Text = ""
   ' ********************************
   ' renumérotation 1ère colonne SENL
   gridSENL.Col = 0
   For i% = 1 To NbEq%
      gridSENL.Row = i%
      gridSENL.Text = "F" & i% & "="
   Next i%
   ' ********************************
   ' renumérotation 1ère ligne Upar
   gridUpar.Row = 0
   gridUpar.Col = 1
   gridUpar.Text = ""
   ' ********************************
   ' renumérotation 1ère colonne Upar
   gridUpar.Col = 0
   For i% = 1 To NbEq%
      gridUpar.Row = i%
      gridUpar.Text = "U(" & i% & ") ="
   Next i%
   ' ********************************
End Sub



Public Sub SolSysNonLin()
   '----------------------------------------------------------
   ' Résolution d'un système non linéaire
   ' de NbEq% équations à NbEq% inconnues
   ' ---------------------------------------------------------
   ' Système : F1[UApar(1),UApar(2),...,UApar(n)]=0
   '           F2[UApar(1),UApar(2),...,UApar(n)]=0
   '           ..................
   '           Fn[UApar(1),UApar(2),...,UApar(n)]=0
   '
   ' On cherche les Upar(i) = UApar(i) + Uvec(i)
   ' qui améliorent le système.
   '
   ' On pose :
   '           Mmat(i,j) = dFi/dUj
   '           Vvec(i)   = - Fi
   '
   ' Equations :  Mmat(i,j) . Uvec(j) = Vvec(i)
   ' On cherche les Uvec(j);
   ' ---------------------------------------------------------
   On Error GoTo Traite_ErreursSysNonLin
   '--------------- Calcul des Mmat(i,j) et Vvec(j) ----------
   ' Initialisations
   OrdreMat% = NbEq%
   NbPar% = NbEq%
   ' ********************************************
   '                ReDims
   ' ********************************************
   Erase SENL$, Upar
   ReDim Mmat(1 To NbEq%, 1 To NbEq%)
   ReDim Upar(1 To NbEq%)
   ReDim UApar(1 To NbEq%)
   ReDim Uvec(1 To NbEq%)
   ReDim Vvec(1 To NbEq%)
   ReDim SENL$(1 To NbEq%)
   ' *************************************************
   '               Initialisations
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' *************************************************
   ' remise en place de la grille de Upar
   ' **************************************
   gridUpar.Rows = NbEq% + 1
   '***************************************
   gridUpar.FixedAlignment(0) = 2
   gridUpar.FixedAlignment(1) = 2
   gridUpar.ColWidth(1) = 1000
   ' ********************************
   ' renumérotation 1ère ligne Upar
   gridUpar.Row = 0
   gridUpar.Col = 1
   gridUpar.Text = ""
   ' ********************************
   ' renumérotation 1ère colonne Upar
   gridUpar.Col = 0
   For i% = 1 To NbEq%
      gridUpar.Row = i%
      gridUpar.Text = "U(" & i% & ") ="
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de SENL$
   ' *************************************************
   gridSENL.Col = 1
   For i% = 1 To NbEq%
      gridSENL.Row = i%
      If gridSENL.Text = "" Then
         gridSENL.Text = ""
         SENL$(i%) = ""
      Else
         SENL$(i%) = gridSENL.Text
      End If
   Next i%
   ' ===============
   ' calcul de Upar
   ' ===============
   ' *************************************************
   ' initialisations
   ' *************************************************
   eps = 0.001
   NbAfMax% = 100
   NbEssaiMax% = 2
   NbEssai% = 0
   Converge = False
   ' *************************************************
   ' Boucle d'essais de calcul des Upar.
   ' Si le premier essai donne un système indéterminé,
   ' on prend d'autres valeurs de départ pour les Upar
   ' *************************************************
   Do
      NbEssai% = NbEssai% + 1
      ' *************************************************
      ' initialisations
      ' *************************************************
      Select Case NbEssai%
      Case 1
         For i% = 1 To NbEq%
            UApar(i%) = 0
            Upar(i%) = 0
         Next i%
      Case 2
         For i% = 1 To NbEq%
            UApar(i%) = i%
            Upar(i%) = i%
         Next i%
      End Select
      NbAf% = 0
      For i% = 1 To NbEq%
         Vvec(i%) = 0
         For j% = 1 To NbEq%
            Mmat(i%, j%) = 0
         Next j%
      Next i%
      ' ********************************************
      ' Boucle d'affinement
      ' ********************************************
      Do
         NbAf% = NbAf% + 1
         For i% = 1 To NbEq%
            ForVar2$ = SENL$(i%)
            Call Parametre(ForVar2$, ParVar2$)
            Call Traitement(ParVar2$, SorVar2$)
            FUj = CSng(SorVar2$)
            ' ++++++++++++++++++++++++++++++++++++++++
            ' ++++++++++  Vvec(i) = -Fi  +++++++++++++
            ' ++++++++++++++++++++++++++++++++++++++++
            Vvec(i%) = -FUj
            For j% = 1 To NbEq%
               ' ++++++++++++++++++++++++++++++++++++++++
               ' ++++++++++  Mmat(i, j) = dFi/dUj  ++++++
               ' ++++++++++++++++++++++++++++++++++++++++
               Upar(j%) = UApar(j%) + eps
               Call Parametre(ForVar2$, ParVar2$)
               Call Traitement(ParVar2$, SorVar2$)
               FUjp = CSng(SorVar2$)
               Mmat(i%, j%) = (FUjp - FUj) / eps
               Upar(j%) = UApar(j%)
            Next j%
         Next i%
         ' -----------------------------
         ' Calcul des Uvec(i)
         ' -----------------------------
         ' Mmat(i,j) . Uvec(j) = Vvec(i) ==> Uvec(i) = Wmat(i,j) . Vvec(j)
         Erreur = False
         Call TrianMat
         If Erreur = True Then
            If NbEssai% = NbEssaiMax% Then
               Message$ = "Système indéterminé"
               MsgBox Message$, 48
               Exit Sub
            Else
               Exit Do
            End If
         End If
         ' -----------------------------
         ' Calcul des Upar(i)
         ' -----------------------------
         For i% = 1 To NbEq%
            Upar(i%) = UApar(i%) + Uvec(i%)
            UApar(i%) = Upar(i%)
         Next i%
         ' -----------------------------
         ' Test de convergence
         ' -----------------------------
         ' Calcul de UvecMax :
         UvecMax = 0
         For i% = 1 To NbEq%
            AbUv = Abs(Uvec(i%))
            If AbUv > UvecMax Then
               UvecMax = AbUv
            End If
         Next i%
         ' Test :
         If UvecMax < eps Then
            ' Convergence
            Converge = True
            Exit Do
         End If
         ' -----------------------------
         ' Test de nombre de boucles
         ' -----------------------------
         If NbAf% = NbAfMax% Then
            If NbEssai% = NbEssaiMax% Then
               Message$ = "Pas de convergence après" & NbAfMax% & "itérations."
               MsgBox Message$, 48
               Exit Sub
            Else
               Exit Do
            End If
         End If
         ' -----------------------------
      Loop
      If Converge = True Then
         Exit Do
      End If
   Loop
   ' *************************************************
   ' affichage des éléments de Upar
   ' *************************************************
   gridUpar.Col = 1
   For i% = 1 To NbEq%
      gridUpar.Row = i%
      gridUpar.Text = Format(Upar(i%), "0.000")
   Next i%
   ' *************************************************
   lblInfo.Caption = ""
   ' *************************************************
   Exit Sub
Traite_ErreursSysNonLin:
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
