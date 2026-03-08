VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreTriNombres 
   Caption         =   "Tri de nombres"
   ClientHeight    =   5760
   ClientLeft      =   1815
   ClientTop       =   1125
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   5730
   Begin MSFlexGridLib.MSFlexGrid GridNbTries 
      Height          =   4335
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridNbATrier 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7646
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
   Begin VB.TextBox txtNbNb 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblNbTries 
      Caption         =   "Nombres triés :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblNbNb 
      Caption         =   "Nombre de nombres à trier :"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblNbATrier 
      Caption         =   "Nombres à trier :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu mnuEnreg 
         Caption         =   "&Enregistrer..."
         Begin VB.Menu mnuEnregNbATrier 
            Caption         =   "&Nombres à trier"
         End
         Begin VB.Menu mnuEnregNbTries 
            Caption         =   "No&mbres triés"
         End
      End
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer..."
      Begin VB.Menu mnuTriParIns 
         Caption         =   "&Tri par insertion"
      End
      Begin VB.Menu mnuTriRapide 
         Caption         =   "&Tri-rapide"
      End
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreTriNombres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NbNombres%
Dim NomTri$
Dim NbATrier()
Dim NbTries()








Private Sub Form_Load()
   '***************************************
   lblInfo.ForeColor = ROUGE
   NbNombres% = 16
   '***************************************
   ReDim NbATrier(1 To NbNombres%)
   ReDim NbTries(1 To NbNombres%)
   ' ********************************
   ' mise en place des grilles
   ' ********************************
   txtNbNb.Text = Format(NbNombres%, "0")
   GridNbATrier.Rows = NbNombres% + 1
   GridNbATrier.Cols = 2
   GridNbTries.Rows = NbNombres% + 1
   GridNbTries.Cols = 2
   ' ********************************
   ' largeur des colonnes et
   ' mentions dans les 1ères lignes
   ' ********************************
   GridNbATrier.FixedAlignment(0) = 2
   GridNbATrier.FixedAlignment(1) = 2
   GridNbATrier.ColWidth(0) = 800
   GridNbATrier.ColWidth(1) = 1600
   GridNbATrier.Row = 0
   GridNbATrier.Col = 0
   GridNbATrier.Text = ""
   GridNbATrier.Col = 1
   GridNbATrier.Text = ""
   '***************************************
   GridNbTries.FixedAlignment(0) = 2
   GridNbTries.FixedAlignment(1) = 2
   GridNbTries.ColWidth(0) = 800
   GridNbTries.ColWidth(1) = 1600
   GridNbTries.Row = 0
   GridNbTries.Col = 0
   GridNbTries.Text = ""
   GridNbTries.Col = 1
   GridNbTries.Text = ""
   '***************************************
   ' ********************************
   ' numérotation 1ères colonnes
   ' ********************************
   GridNbATrier.Col = 0
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      GridNbATrier.Text = Format(i%, "0")
   Next i%
   '************************************
   GridNbTries.Col = 0
   For i% = 1 To NbNombres%
      GridNbTries.Row = i%
      GridNbTries.Text = Format(i%, "0")
   Next i%
   '************************************
   ' ********************************
   '       Valeurs par défaut
   ' ********************************
   NbATrier(1) = 15
   NbATrier(2) = 13
   NbATrier(3) = 7
   NbATrier(4) = 12
   NbATrier(5) = 1
   NbATrier(6) = 9
   NbATrier(7) = 8
   NbATrier(8) = 10
   NbATrier(9) = 4
   NbATrier(10) = 16
   NbATrier(11) = 5
   NbATrier(12) = 3
   NbATrier(13) = 6
   NbATrier(14) = 2
   NbATrier(15) = 11
   NbATrier(16) = 14
   ' ********************************
   ' placement des valeurs par défaut
   ' ********************************
   GridNbATrier.Col = 1
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      GridNbATrier.Text = Format(NbATrier(i%), "0.000")
   Next i%
   ' ****************************************
End Sub

Private Sub ProduitPol()
   DegPPpol% = DegPpol% + DegP2pol%
   ' ********************************************
   '               ReDims
   ' ********************************************
   ReDim Ppol(0 To DegPpol%)
   ReDim P2pol(0 To DegP2pol%)
   ReDim PPpol(0 To DegPPpol%)
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' **************************************
   ' ***** mise en place des grilles *****
   ' **************************************
   FenetrePolynome.lblValDegPP.Caption = Format(DegPPpol%, "0")
   FenetrePolynome.GridPPpol.Rows = DegPPpol% + 2
   FenetrePolynome.GridPPpol.Cols = 2
   ' ********************************
   ' numérotation 1ère colonnes
   ' ********************************
   FenetrePolynome.GridPPpol.Col = 0
   For i% = 0 To DegPPpol%
      FenetrePolynome.GridPPpol.Row = i% + 1
      FenetrePolynome.GridPPpol.Text = Format(i%, "0")
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.GridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.GridPpol.Row = i% + 1
      If FenetrePolynome.GridPpol.Text = "" Then
         FenetrePolynome.GridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.GridPpol.Text)
      End If
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de P2pol
   ' *************************************************
   FenetrePolynome.GridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.GridP2pol.Row = i% + 1
      If FenetrePolynome.GridP2pol.Text = "" Then
         FenetrePolynome.GridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.GridP2pol.Text)
      End If
   Next i%
   ' *************************************************
   ' calcul de PP produit de P et P2
   ' *************************************************
   For k% = 0 To DegPPpol%
      PPpol(k%) = 0
   Next k%
   For i% = 0 To DegPpol%
      For j% = 0 To DegP2pol%
         k% = i% + j%
         PPpol(k%) = PPpol(k%) + Ppol(i%) * P2pol(j%)
      Next j%
   Next i%
   ' *************************************************
   ' affichage des éléments de PPpol
   ' *************************************************
   FenetrePolynome.GridPPpol.Col = 1
   For i% = 0 To DegPPpol%
      FenetrePolynome.GridPPpol.Row = i% + 1
      FenetrePolynome.GridPPpol.Text = Format(PPpol(i%), "0.000")
   Next i%
   ' ********************************************
   lblInfo.Caption = ""
   ' ********************************************
End Sub

Private Sub QuotientPol()
   If DegPpol% < DegP2pol% Then
      MsgBox "le degré de P doit être supérieur ou égal à celui de P2 !", 48, "POLYNOME"
      Exit Sub
   End If
   DegQpol% = DegPpol% - DegP2pol%
   If DegP2pol% > 0 Then
      DegRpol% = DegP2pol% - 1
   Else
      DegRpol% = 0
   End If
   ' ***************************************
   '               ReDims
   ' ********************************************
   ReDim Ppol(0 To DegPpol%)
   ReDim P2pol(0 To DegP2pol%)
   ReDim Qpol(0 To DegQpol%)
   ReDim Rpol(0 To DegPpol%)
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' **************************************
   ' ***** mise en place des grilles ****
   ' **************************************
   FenetrePolynome.lblValDegQ.Caption = Format(DegQpol%, "0")
   FenetrePolynome.GridQpol.Rows = DegQpol% + 2
   FenetrePolynome.GridQpol.Cols = 2
   FenetrePolynome.lblValDegR.Caption = Format(DegRpol%, "0")
   FenetrePolynome.GridRpol.Rows = DegRpol% + 2
   FenetrePolynome.GridRpol.Cols = 2
   ' ********************************
   ' numérotation 1ères colonnes
   ' ********************************
   FenetrePolynome.GridQpol.Col = 0
   For i% = 0 To DegQpol%
      FenetrePolynome.GridQpol.Row = i% + 1
      FenetrePolynome.GridQpol.Text = Format(i%, "0")
   Next i%
   FenetrePolynome.GridRpol.Col = 0
   For i% = 0 To DegRpol%
      FenetrePolynome.GridRpol.Row = i% + 1
      FenetrePolynome.GridRpol.Text = Format(i%, "0")
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.GridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.GridPpol.Row = i% + 1
      If FenetrePolynome.GridPpol.Text = "" Then
         FenetrePolynome.GridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.GridPpol.Text)
      End If
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de P2pol
   ' *************************************************
   FenetrePolynome.GridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.GridP2pol.Row = i% + 1
      If FenetrePolynome.GridP2pol.Text = "" Then
         FenetrePolynome.GridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.GridP2pol.Text)
      End If
   Next i%
   ' *************************************************
   ' calcul de Q quotient de P par P2
   ' et R reste
   ' *************************************************
   For k% = 0 To DegQpol%
      Qpol(k%) = 0
   Next k%
   For k% = 0 To DegPpol%
      Rpol(k%) = Ppol(k%)
   Next k%
   For i% = DegPpol% To DegP2pol% Step -1
      Qpol(i% - DegP2pol%) = Rpol(i%) / P2pol(DegP2pol%)
      For j% = i% To i% - DegP2pol% Step -1
         Rpol(j%) = Rpol(j%) - Qpol(i% - DegP2pol%) * P2pol(j% - i% + DegP2pol%)
      Next j%
   Next i%
   ' *************************************************
   ' affichage des éléments de Qpol et Rpol
   ' *************************************************
   FenetrePolynome.GridQpol.Col = 1
   For i% = 0 To DegQpol%
      FenetrePolynome.GridQpol.Row = i% + 1
      FenetrePolynome.GridQpol.Text = Format(Qpol(i%), "0.000")
   Next i%
   FenetrePolynome.GridRpol.Col = 1
   For i% = 0 To DegRpol%
      FenetrePolynome.GridRpol.Row = i% + 1
      FenetrePolynome.GridRpol.Text = Format(Rpol(i%), "0.000")
   Next i%
   ' ********************************************
   lblInfo.Caption = ""
   ' ********************************************
End Sub



















Private Sub mnuEnregP_Click()
   NomPol$ = "P"
   Call EnregPol
End Sub


Private Sub mnuEnregP2_Click()
   NomPol$ = "P2"
   Call EnregPol
End Sub



Private Sub mnuEnregPP_Click()
   NomPol$ = "PP"
   Call EnregPol
End Sub



Private Sub mnuEnregQ_Click()
   NomPol$ = "Q"
   Call EnregPol
End Sub



Private Sub mnuEnregR_Click()
   NomPol$ = "R"
   Call EnregPol
End Sub



Private Sub mnuOuvrirP_Click()
   NomPol$ = "P"
   Call OuvrePol
End Sub



Private Sub mnuOuvrirP2_Click()
   NomPol$ = "P2"
   Call OuvrePol
End Sub


Private Sub gridNbATrier_KeyPress(KeyAscii As Integer)
   EleNbText$ = GridNbATrier.Text
   Select Case KeyAscii
   Case 32 To 168
      EleNbCar$ = Chr(KeyAscii)
      EleNbText$ = EleNbText$ & EleNbCar$
      GridNbATrier.Text = EleNbText$
   Case 8
      If Len(GridNbATrier.Text) > 0 Then
         EleNbText$ = Left$(EleNbText$, Len(EleNbText$) - 1)
         GridNbATrier.Text = EleNbText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub mnuEnregNbATrier_Click()
   On Error GoTo Traite_ErreursEnregNbATrier
   Maths.ctrlCMDialog.DefaultExt = "nbs"
   Maths.ctrlCMDialog.Filter = "Nombres (*.nbs)|*.nbs"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   '-----------------------------------------------------
   ' Création du fichier de nombres
   ' et écriture de ces éléments dans le fichier
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbNombres%
   GridNbATrier.Col = 1
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      Write #1, GridNbATrier.Text
   Next i%
   Close #1
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnregNbATrier:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48, "EnregNbATrier"
   End Select
   Exit Sub
End Sub


Private Sub mnuEnregNbTries_Click()
   On Error GoTo Traite_ErreursEnregNbTries
   Maths.ctrlCMDialog.DefaultExt = "nbs"
   Maths.ctrlCMDialog.Filter = "Nombres (*.nbs)|*.nbs"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   '-----------------------------------------------------
   ' Création du fichier de nombres
   ' et écriture de ces éléments dans le fichier
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Write #1, NbNombres%
   GridNbTries.Col = 1
   For i% = 1 To NbNombres%
      GridNbTries.Row = i%
      Write #1, GridNbTries.Text
   Next i%
   Close #1
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnregNbTries:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48, "EnregNbTries"
   End Select
   Exit Sub
End Sub
Private Sub mnuOuvrir_Click()
   '-----------------------------------------------------
   'On Error GoTo Traite_ErreursOuvNomb
   Maths.ctrlCMDialog.Filter = "Nombres (*.nbs)|*.nbs"
   ' nom de fichier et chemin doivent exister
   ' sinon apparait un message d'erreur spécifique
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   '-----------------------------------------------------
   ' Ouverture et lecture du fichier
   ' de nombres
   ' et écriture de ces coefficients
   ' dans NbATrier(i%)
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, NbNombres%
   ' *************************************************
   ' placement du nombre de nombres, ce qui provoque
   ' le redimentionnement des grilles
   ' *************************************************
   txtNbNb.Text = Format(NbNombres%, "0")
   ' *************************************************
   ' lecture des nombres
   ' *************************************************
   For i% = 1 To NbNombres%
      Input #1, NbATrier(i%)
   Next i%
   Close #1
   ' ***********************************************
   ' placement des nouveaux nombres
   ' ***********************************************
   GridNbATrier.Col = 1
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      GridNbATrier.Text = Format(NbATrier(i%), "0.000")
   Next i%
   ' ****************************************
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursOuvNomb:
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
   FenetreTriNombres.Hide
End Sub


Private Sub mnuQuotient_Click()
   Call QuotientPol
End Sub


Public Sub OuvrePol()
   '-----------------------------------------------------
   On Error GoTo Traite_ErreursOuvPol
   Maths.ctrlCMDialog.Filter = "Polynôme (*.pol)|*.pol"
   ' nom de fichier et chemin doivent exister
   ' sinon apparait un message d'erreur spécifique
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   '-----------------------------------------------------
   ' Ouverture et lecture du fichier
   ' coefficients d'un polynôme
   ' et écriture de ces coefficients
   ' dans Ppol(i%) ou P2pol(i%)
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Input #1, DegPolLoc%
   ' *************************************************
   ' placement du degré du polynôme, ce qui provoque
   ' le redimentionnement des grilles
   ' *************************************************
   If NomPol$ = "P" Then
      txtValDegP.Text = Format(DegPolLoc%, "0")
   ElseIf NomPol$ = "P2" Then
      txtValDegP2.Text = Format(DegPolLoc%, "0")
   End If
   ' *************************************************
   ' lecture des coefficients du polynôme
   ' *************************************************
   For i% = 0 To DegPolLoc%
      If NomPol$ = "P" Then
         Input #1, Ppol(i%)
      ElseIf NomPol$ = "P2" Then
         Input #1, P2pol(i%)
      End If
   Next i%
   Close #1
   ' ***********************************************
   ' placement des nouveaux coefficients du polynôme
   ' ***********************************************
   If NomPol$ = "P" Then
      FenetrePolynome.GridPpol.Col = 1
      For i% = 0 To DegPolLoc%
         FenetrePolynome.GridPpol.Row = i% + 1
         FenetrePolynome.GridPpol.Text = Format(Ppol(i%), "0.000")
      Next i%
   ElseIf NomPol$ = "P2" Then
      FenetrePolynome.GridP2pol.Col = 1
      For i% = 0 To DegPolLoc%
         FenetrePolynome.GridP2pol.Row = i% + 1
         FenetrePolynome.GridP2pol.Text = Format(P2pol(i%), "0.000")
      Next i%
   End If
   ' ****************************************
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursOuvPol:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Public Sub EnregPol()
   On Error GoTo Traite_ErreursEnregPol
   Maths.ctrlCMDialog.DefaultExt = "pol"
   Maths.ctrlCMDialog.Filter = "Polynôme (*.pol)|*.pol"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Ppol, P2pol, PPpol, Qpol, Rpol
   ReDim Ppol(0 To DegPpol%)
   ReDim P2pol(0 To DegP2pol%)
   ReDim PPpol(0 To DegPPpol%)
   ReDim Qpol(0 To DegQpol%)
   ReDim Rpol(0 To DegRpol%)
   ' ***********************************************************
   ' affectation de leurs valeurs aux coefficients des polynômes
   ' ***********************************************************
   FenetrePolynome.GridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.GridPpol.Row = i% + 1
      If FenetrePolynome.GridPpol.Text = "" Then
         FenetrePolynome.GridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.GridPpol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.GridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.GridP2pol.Row = i% + 1
      If FenetrePolynome.GridP2pol.Text = "" Then
         FenetrePolynome.GridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.GridP2pol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.GridPPpol.Col = 1
   For i% = 0 To DegPPpol%
      FenetrePolynome.GridPPpol.Row = i% + 1
      If FenetrePolynome.GridPPpol.Text = "" Then
         FenetrePolynome.GridPPpol.Text = 0
         PPpol(i%) = 0
      Else
         PPpol(i%) = CSng(FenetrePolynome.GridPPpol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.GridQpol.Col = 1
   For i% = 0 To DegQpol%
      FenetrePolynome.GridQpol.Row = i% + 1
      If FenetrePolynome.GridQpol.Text = "" Then
         FenetrePolynome.GridQpol.Text = "0"
         Qpol(i%) = 0
      Else
         Qpol(i%) = CSng(FenetrePolynome.GridQpol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.GridRpol.Col = 1
   For i% = 0 To DegRpol%
      FenetrePolynome.GridRpol.Row = i% + 1
      If FenetrePolynome.GridRpol.Text = "" Then
         FenetrePolynome.GridRpol.Text = "0"
         Rpol(i%) = 0
      Else
         Rpol(i%) = CSng(FenetrePolynome.GridRpol.Text)
      End If
   Next i%
   ' *************************************************
   '-----------------------------------------------------
   ' Création du fichier de coefficients du polynôme
   ' et écriture de ces éléments dans le fichier
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   If NomPol$ = "P" Then
      Write #1, DegPpol%
      For i% = 0 To DegPpol%
         Write #1, Ppol(i%)
      Next i%
   ElseIf NomPol$ = "P2" Then
      Write #1, DegP2pol%
      For i% = 0 To DegP2pol%
         Write #1, P2pol(i%)
      Next i%
   ElseIf NomPol$ = "PP" Then
      Write #1, DegPPpol%
      For i% = 0 To DegPPpol%
         Write #1, PPpol(i%)
      Next i%
   ElseIf NomPol$ = "Q" Then
      Write #1, DegQpol%
      For i% = 0 To DegQpol%
         Write #1, Qpol(i%)
      Next i%
   ElseIf NomPol$ = "R" Then
      Write #1, DegRpol%
      For i% = 0 To DegRpol%
         Write #1, Rpol(i%)
      Next i%
   End If
   Close #1
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnregPol:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48, "EnregPol"
   End Select
   Exit Sub
End Sub

Public Sub RacinesExactes()
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.GridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.GridPpol.Row = i% + 1
      If FenetrePolynome.GridPpol.Text = "" Then
         FenetrePolynome.GridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.GridPpol.Text)
      End If
   Next i%
   ' *************************************************
   Select Case DegPpol%
   Case Is < 1
      Beep
      MsgBox "le degré de P doit être supérieur ou égal à 1 !", 48, "POLYNOME"
      Exit Sub
   Case 1
      If Ppol(1) = 0 Then
         Beep
         MsgBox "le polynôme P est de degré zéro !", 48, "POLYNOME"
         Exit Sub
      End If
      '*****************************************************************
      '*************  Racines d'un polynôme de degré 1  ****************
      '*****************************************************************
      XPol1 = -Ppol(0) / Ppol(1)
      TexteRacines$ = " Une racine :"
      lblValRacines.Caption = TexteRacines$
      FenetrePolynome.GridRacines.Row = 1
      FenetrePolynome.GridRacines.Col = 1
      FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
      FenetrePolynome.GridRacines.Col = 2
      FenetrePolynome.GridRacines.Text = "0"
   Case 2
      If Ppol(2) = 0 Then
         Beep
         MsgBox "le polynôme P est de degré inférieur à 2 !", 48, "POLYNOME"
         Exit Sub
      End If
      '*****************************************************************
      '*************  Racines d'un polynôme de degré 2  ****************
      '*****************************************************************
      Dis = Ppol(1) ^ 2 - 4 * Ppol(2) * Ppol(0)
      ReelPol = -Ppol(1) / Ppol(2) / 2
      If Abs(Dis) < 0.000001 Then
         TexteRacines$ = " Une racine double :"
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         ImPol = RacDis / Ppol(2) / 2
         XPol1 = ReelPol - ImPol
         XPol2 = ReelPol + ImPol
         TexteRacines$ = " Deux racines réelles :"
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol2, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      Else
         RacDis = Sqr(-Dis)
         ImPol = Abs(RacDis / Ppol(2) / 2)
         TexteRacines$ = " Deux racines complexes conjuguées :"
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(-ImPol, "0.00")
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(ImPol, "0.00")
      End If
      lblValRacines.Caption = TexteRacines$
   Case 3
      If Ppol(3) = 0 Then
         Beep
         MsgBox "le polynôme P est de degré inférieur à 3 !", 48, "POLYNOME"
         Exit Sub
      End If
      '*****************************************************************
      '*************  Racines d'un polynôme de degré 3  ****************
      '*************          Méthode de CARDAN         ****************
      '*****************************************************************
      TexteRacines$ = " Méthode de CARDAN; "
      Ploc = Ppol(1) / Ppol(3) / 3 - Ppol(2) * Ppol(2) / Ppol(3) / Ppol(3) / 9
      Qloc = Ppol(2) * Ppol(2) * Ppol(2) / Ppol(3) / Ppol(3) / Ppol(3) / 27 - Ppol(2) * Ppol(1) / Ppol(3) / Ppol(3) / 6 + Ppol(0) / Ppol(3) / 2
      Dis = Qloc * Qloc + Ploc * Ploc * Ploc
      If Abs(Dis) < 0.000001 Then
         If Qloc = 0 Then
            ReelPol = -Ppol(2) / Ppol(3) / 3
            TexteRacines$ = TexteRacines$ & " Une racine triple :"
            FenetrePolynome.GridRacines.Row = 1
            FenetrePolynome.GridRacines.Col = 1
            FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
            FenetrePolynome.GridRacines.Col = 2
            FenetrePolynome.GridRacines.Text = "0"
            FenetrePolynome.GridRacines.Row = 2
            FenetrePolynome.GridRacines.Col = 1
            FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
            FenetrePolynome.GridRacines.Col = 2
            FenetrePolynome.GridRacines.Text = "0"
            FenetrePolynome.GridRacines.Row = 3
            FenetrePolynome.GridRacines.Col = 1
            FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
            FenetrePolynome.GridRacines.Col = 2
            FenetrePolynome.GridRacines.Text = "0"
         Else
            XPol1 = 2 * Qloc / Ploc - Ppol(2) / Ppol(3) / 3
            XPol2 = -Qloc / Ploc - Ppol(2) / Ppol(3) / 3
            TexteRacines$ = TexteRacines$ & " Deux racines réelles : une simple et une double :"
            FenetrePolynome.GridRacines.Row = 1
            FenetrePolynome.GridRacines.Col = 1
            FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
            FenetrePolynome.GridRacines.Col = 2
            FenetrePolynome.GridRacines.Text = "0"
            FenetrePolynome.GridRacines.Row = 2
            FenetrePolynome.GridRacines.Col = 1
            FenetrePolynome.GridRacines.Text = Format(XPol2, "0.00")
            FenetrePolynome.GridRacines.Col = 2
            FenetrePolynome.GridRacines.Text = "0"
            FenetrePolynome.GridRacines.Row = 3
            FenetrePolynome.GridRacines.Col = 1
            FenetrePolynome.GridRacines.Text = Format(XPol2, "0.00")
            FenetrePolynome.GridRacines.Col = 2
            FenetrePolynome.GridRacines.Text = "0"
         End If
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         GAloc = Sgn(-Qloc + RacDis) * Abs(-Qloc + RacDis) ^ (1 / 3)
         GBloc = Sgn(-Qloc - RacDis) * Abs(-Qloc - RacDis) ^ (1 / 3)
         ReelPol1 = GAloc + GBloc - Ppol(2) / 3 / Ppol(3)
         ReelPol2 = (-GAloc - GBloc) / 2 - Ppol(2) / 3 / Ppol(3)
         ImPol = Abs(Sqr(3) / 2 * (GAloc - GBloc))
         TexteRacines$ = TexteRacines$ & " Trois racines : une réelle et deux complexes conjuguées"
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol2, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(-ImPol, "0.00")
         FenetrePolynome.GridRacines.Row = 3
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol2, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(ImPol, "0.00")
      Else
         RacDis = Sqr(-Dis)
         RacPloc = Sqr(-Ploc)
         If Qloc = 0 Then
            Phi = PI / 6
         Else
            '******** Routine Phi = arcos(XT)/3 *******
            XTloc = -Qloc / RacPloc / RacPloc / RacPloc
            LMloc = 0
            LSloc = PI / 3
            Do
               Phi = LMloc / 2 + LSloc / 2
               CTloc = Cos(3 * Phi)
               If Abs(XTloc - CTloc) < 0.000001 Then Exit Do
               If XTloc > CTloc Then LSloc = Phi Else LMloc = Phi
            Loop
            '******************************************
         End If
         GPloc = RacPloc * Cos(Phi)
         GQloc = RacPloc * Sqr(3) * Sin(Phi)
         XPol1 = 2 * GPloc - Ppol(2) / 3 / Ppol(3)
         XPol2 = -GPloc - GQloc - Ppol(2) / 3 / Ppol(3)
         XPol3 = -GPloc + GQloc - Ppol(2) / 3 / Ppol(3)
         TexteRacines$ = TexteRacines$ & " Trois racines réelles :"
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol2, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 3
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol3, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      End If
      lblValRacines.Caption = TexteRacines$
   Case 4
      If Ppol(4) = 0 Then
         Beep
         MsgBox "le polynôme P est de degré inférieur à 4 !", 48, "POLYNOME"
         Exit Sub
      End If
      '*****************************************************************
      '*************  Racines d'un polynôme de degré 4  ****************
      '*************          Méthode de FERRARI        ****************
      '*****************************************************************
      TexteRacines$ = "      Méthode de FERRARI"
      A3loc = 2
      A2loc = -Ppol(2) / Ppol(4)
      A1loc = Ppol(3) * Ppol(1) / Ppol(4) / Ppol(4) / 2 - 2 * Ppol(0) / Ppol(4)
      A0loc = -Ppol(3) * Ppol(3) * Ppol(0) / Ppol(4) / Ppol(4) / Ppol(4) / 4 + Ppol(2) * Ppol(0) / Ppol(4) / Ppol(4) - Ppol(1) * Ppol(1) / Ppol(4) / Ppol(4) / 4
      Ploc = A1loc / A3loc / 3 - A2loc * A2loc / A3loc / A3loc / 9
      Qloc = A2loc * A2loc * A2loc / A3loc / A3loc / A3loc / 27 - A2loc * A1loc / A3loc / A3loc / 6 + A0loc / A3loc / 2
      Dis = Qloc * Qloc + Ploc * Ploc * Ploc
      If Abs(Dis) < 0.000001 Then
         If Qloc = 0 Then
            Zloc = -A2loc / A3loc / 3
            APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
            If APloc < 0 Then
               MsgBox "erreur : quantité négative !", 48, "POLYNOME"
               Exit Sub
            End If
            BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
            If BPloc < 0 Then
               MsgBox "erreur : quantité négative !", 48, "POLYNOME"
               Exit Sub
            End If
         Else
            Zloc = 2 * Qloc / Ploc - A2loc / A3loc / 3
            APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
            If APloc >= 0 Then
               BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
            End If
            If APloc < 0 Or BPloc < 0 Then
               Zloc = -Qloc / Ploc - A2loc / A3loc / 3
               APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
               If APloc < 0 Then
                  MsgBox "erreur : quantité négative !", 48, "POLYNOME"
                  Exit Sub
               End If
               BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
               If BPloc < 0 Then
                  MsgBox "erreur : quantité négative !", 48, "POLYNOME"
                  Exit Sub
               End If
            End If
         End If
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         GAloc = Sgn(-Qloc + RacDis) * Abs(-Qloc + RacDis) ^ (1 / 3)
         GBloc = Sgn(-Qloc - RacDis) * Abs(-Qloc - RacDis) ^ (1 / 3)
         Zloc = GAloc + GBloc - A2loc / 3 / A3loc
         APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
         If APloc < 0 Then
            MsgBox "erreur : quantité négative !", 48, "POLYNOME"
            Exit Sub
         End If
         BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
         If BPloc < 0 Then
            MsgBox "erreur : quantité négative !", 48, "POLYNOME"
            Exit Sub
         End If
      Else
         RacDis = Sqr(-Dis)
         RacPloc = Sqr(-Ploc)
         If Qloc = 0 Then
            Phi = PI / 6
         Else
            '******** Routine Phi = arcos(XT)/3 *******
            XTloc = -Qloc / RacPloc / RacPloc / RacPloc
            LMloc = 0
            LSloc = PI / 3
            Do
               Phi = LMloc / 2 + LSloc / 2
               CTloc = Cos(3 * Phi)
               If Abs(XTloc - CTloc) < 0.000001 Then Exit Do
               If XTloc > CTloc Then LSloc = Phi Else LMloc = Phi
            Loop
            '******************************************
         End If
         GPloc = RacPloc * Cos(Phi)
         GQloc = RacPloc * Sqr(3) * Sin(Phi)
         Zloc = 2 * GPloc - A2loc / 3 / A3loc
         APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
         If APloc >= 0 Then
            BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
         End If
         If APloc < 0 Or BPloc < 0 Then
            Zloc = -GPloc - GQloc - A2loc / 3 / A3loc
            APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
            If APloc >= 0 Then
               BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
            End If
            If APloc < 0 Or BPloc < 0 Then
               Zloc = -GPloc + GQloc - A2loc / 3 / A3loc
               APloc = Ppol(3) * Ppol(3) / Ppol(4) / Ppol(4) / 4 - Ppol(2) / Ppol(4) + 2 * Zloc
               If APloc < 0 Then
                  MsgBox "erreur : quantité négative !", 48, "POLYNOME"
                  Exit Sub
               End If
               BPloc = Zloc * Zloc - Ppol(0) / Ppol(4)
               If BPloc < 0 Then
                  MsgBox "erreur : quantité négative !", 48, "POLYNOME"
                  Exit Sub
               End If
            End If
         End If
      End If
      APloc = Sqr(APloc)
      BPloc = Sqr(BPloc)
      CPloc = Ppol(3) * Zloc / Ppol(4) - Ppol(1) / Ppol(4)
      If CPloc < 0 Then BPloc = -BPloc
      BAloc = Ppol(3) / Ppol(4) / 2 + APloc
      CAloc = Zloc + BPloc
      DMloc = BAloc * BAloc - 4 * CAloc
      ReelPol = -BAloc / 2
      If Abs(DMloc) < 0.000001 Then
         XPol1 = ReelPol
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      ElseIf DMloc >= 0 Then
         ImPol = Sqr(DMloc) / 2
         XPol2 = ReelPol - ImPol
         XPol3 = ReelPol + ImPol
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol2, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol3, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      Else
         ImPol = Sqr(-DMloc) / 2
         FenetrePolynome.GridRacines.Row = 1
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(-ImPol, "0.00")
         FenetrePolynome.GridRacines.Row = 2
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(ImPol, "0.00")
      End If
      BBloc = Ppol(3) / Ppol(4) / 2 - APloc
      CBloc = Zloc - BPloc
      DMloc = BBloc * BBloc - 4 * CBloc
      ReelPol = -BBloc / 2
      If Abs(DMloc) < 0.000001 Then
         XPol1 = ReelPol
         FenetrePolynome.GridRacines.Row = 3
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 4
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol1, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      ElseIf DMloc >= 0 Then
         ImPol = Sqr(DMloc) / 2
         XPol2 = ReelPol - ImPol
         XPol3 = ReelPol + ImPol
         FenetrePolynome.GridRacines.Row = 3
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol2, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
         FenetrePolynome.GridRacines.Row = 4
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(XPol3, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = "0"
      Else
         ImPol = Sqr(-DMloc) / 2
         FenetrePolynome.GridRacines.Row = 3
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(-ImPol, "0.00")
         FenetrePolynome.GridRacines.Row = 4
         FenetrePolynome.GridRacines.Col = 1
         FenetrePolynome.GridRacines.Text = Format(ReelPol, "0.00")
         FenetrePolynome.GridRacines.Col = 2
         FenetrePolynome.GridRacines.Text = Format(ImPol, "0.00")
      End If
      lblValRacines.Caption = TexteRacines$
   Case Is > 4
      Message$ = "Impossible de calculer de manière exacte"
      Message$ = Message$ & Chr$(13) & "les racines d'un polynôme de degré supérieur à 4 !"
      MsgBox Message$, 48, "POLYNOME"
   End Select
End Sub

Private Sub mnuTriParIns_Click()
   ' *******************************************************
   ' affectation de leurs valeurs aux éléments de NbATrier()
   ' *******************************************************
   GridNbATrier.Col = 1
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      If GridNbATrier.Text = "" Then
         GridNbATrier.Text = "0"
         NbATrier(i%) = 0
      Else
         NbATrier(i%) = CSng(GridNbATrier.Text)
      End If
   Next i%
   ' ********************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   ' ********************************************
   Call TriParInsertionSM(NbNombres%, NbATrier(), NbTries())
   ' *************************************************
   ' affichage des éléments de NbTries()
   ' *************************************************
   GridNbTries.Col = 1
   For i% = 1 To NbNombres%
      GridNbTries.Row = i%
      GridNbTries.Text = Format(NbTries(i%), "0.000")
   Next i%
   ' ********************************************
   lblInfo.Caption = "Tri par insertion"
   ' ********************************************
End Sub
Public Sub TriParInsertionAM(NbDeNb%, NbEnDesordre(), NbEnOrdre(), PermTpi%())
' *****************************************
' *    Tri par insertion (avec mémoire)   *
' *****************************************
' --------------------------------------------------------------------------
' Principe : On prend le premier nombre; on prend le deuxième nombre et on le
' classe en fonction du premier; on prend le troisième nombre et on le classe
' en en fonction des deux premiers; et ainsi de suite.
' Nombre de comparaisons à effectuer pour n nombres : environ n(n-1)/2
' --------------------------------------------------------------------------
' En entrée :     NbDeNb%        =  Nombre de nombres à trier
'                 NbEnDesordre() =  Liste des nombres à trier
'
' En sortie :     NbEnOrdre()    =  Liste des nombres triés
'                 PermTpi%()     =  Permutations effectuées lors du tri
'                                   [PermTpi%(j%) = ancienne position du nombre
'                                    placé après tri en position j%]
' --------------------------------------------------------------------------
If NbDeNb% < 2 Then
   Message$ = " Le nombre de nombres à trier doit être supérieur à 1 !"
   MsgBox Message$, 48, "TriParInsertionAM"
   Exit Sub
End If
' ---------------------------------------------------
' ********************************
' Algorithme de tri
' ********************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
For i% = 2 To NbDeNb%
   NbEnOrdre(i%) = 0
   PermTpi%(i%) = i%
Next i%
' ---------------------------------------------------
NbEnOrdre(1) = NbEnDesordre(1)
PermTpi%(1) = 1
For i% = 2 To NbDeNb%
   NbIns = NbEnDesordre(i%)
   NbEnOrdre(i%) = NbIns
   For j% = 1 To i% - 1
      If NbIns < NbEnOrdre(j%) Then
         For k% = i% - 1 To j% Step -1
            NbEnOrdre(k% + 1) = NbEnOrdre(k%)
            PermTpi%(k% + 1) = PermTpi%(k%)
         Next k%
         NbEnOrdre(j%) = NbIns
         PermTpi%(j%) = i%
         Exit For
      End If
   Next j%
Next i%
' --------------------------------------------------
End Sub

Public Sub TriParInsertionSM(NbDeNb%, NbEnDesordre(), NbEnOrdre())
' *****************************************
' *    Tri par insertion (sans mémoire)   *
' *****************************************
' --------------------------------------------------------------------------
' Principe : On prend le premier nombre; on prend le deuxième nombre et on le
' classe en fonction du premier; on prend le troisième nombre et on le classe
' en en fonction des deux premiers; et ainsi de suite.
' Nombre de comparaisons à effectuer pour n nombres : environ n(n-1)/2
' --------------------------------------------------------------------------
' En entrée :     NbDeNb%        =  Nombre de nombres à trier
'                 NbEnDesordre() =  Liste des nombres à trier
'
' En sortie :     NbEnOrdre()    =  Liste des nombres triés
' --------------------------------------------------------------------------
If NbDeNb% < 2 Then
   Message$ = " Le nombre de nombres à trier doit être supérieur à 1 !"
   MsgBox Message$, 48, "TriParInsertionSM"
   Exit Sub
End If
' ---------------------------------------------------
' ********************************
' Algorithme de tri
' ********************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
For i% = 2 To NbDeNb%
   NbEnOrdre(i%) = 0
Next i%
' ---------------------------------------------------
NbEnOrdre(1) = NbEnDesordre(1)
For i% = 2 To NbDeNb%
   NbIns = NbEnDesordre(i%)
   NbEnOrdre(i%) = NbIns
   For j% = 1 To i% - 1
      If NbIns < NbEnOrdre(j%) Then
         For k% = i% - 1 To j% Step -1
            NbEnOrdre(k% + 1) = NbEnOrdre(k%)
         Next k%
         NbEnOrdre(j%) = NbIns
         Exit For
      End If
   Next j%
Next i%
' --------------------------------------------------
End Sub
Public Sub TriRapideAM(NbDeNb%, NbEnDesordre(), NbEnOrdre(), PermTra%())
' **********************************
' *    Tri-rapide (avec mémoire)   *
' **********************************
' --------------------------------------------------------------------------
' Principe : On prend le premier nombre (appelé pivot); on compare les autres
' nombres au pivot et on les classe en 2 groupes : le premier constitué des
' nombres qui lui sont inférieurs et le second de ceux qui lui sont supérieurs;
' ce tri partiel effectué,on prend alors le premier groupe et on renouvelle
' avec lui la manipulation précedente, de même avec le deuxième groupe, etc...
' Nombre de comparaisons à effectuer pour n nombres : environ n.(log à base 2)(n)
'
' C'est la borne la meilleure que l'on puisse atteindre en n'utilisant que des
' comparaisons.
' --------------------------------------------------------------------------
' En entrée :     NbDeNb%        =  Nombre de nombres à trier
'                 NbEnDesordre() =  Liste des nombres à trier
'
' En sortie :     NbEnOrdre()    =  Liste des nombres triés
'                 PermTra%()     =  Permutations effectuées lors du tri
'                                   [PermTpi%(j%) = ancienne position du nombre
'                                    placé après tri en position j%]
' --------------------------------------------------------------------------
If NbDeNb% < 2 Then
   Message$ = " Le nombre de nombres à trier doit être supérieur à 1 !"
   MsgBox Message$, 48, "TriRapide"
   Exit Sub
End If
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
For i% = 1 To NbDeNb%
   PermTra%(i%) = i%
Next i%
' ---------------------------------------------------
For i% = 1 To NbDeNb%
   NbEnOrdre(i%) = NbEnDesordre(i%)
Next i%
' ---------------------------------------------------
' ********************************
' Algorithme de tri
' ********************************
Call TriPartielAM(1, NbDeNb%, 1, NbEnOrdre(), PermTra%())
' -------------------------------------------------------
End Sub

Public Sub TriRapideSM(NbDeNb%, NbEnDesordre(), NbEnOrdre())
' *********************************
' *    Tri-rapide (sans mémoire)  *
' *********************************
' --------------------------------------------------------------------------
' Principe : On prend le premier nombre (appelé pivot); on compare les autres
' nombres au pivot et on les classe en 2 groupes : le premier constitué des
' nombres qui lui sont inférieurs et le second de ceux qui lui sont supérieurs;
' ce tri partiel effectué,on prend alors le premier groupe et on renouvelle
' avec lui la manipulation précedente, de même avec le deuxième groupe, etc...
' Nombre de comparaisons à effectuer pour n nombres : environ n.(log à base 2)(n)
'
' C'est la borne la meilleure que l'on puisse atteindre en n'utilisant que des
' comparaisons.
' --------------------------------------------------------------------------
' En entrée :     NbDeNb%        =  Nombre de nombres à trier
'                 NbEnDesordre() =  Liste des nombres à trier
'
' En sortie :     NbEnOrdre()    =  Liste des nombres triés
' --------------------------------------------------------------------------
If NbDeNb% < 2 Then
   Message$ = " Le nombre de nombres à trier doit être supérieur à 1 !"
   MsgBox Message$, 48, "TriRapide"
   Exit Sub
End If
' ---------------------------------------------------
' Initialisation
' ---------------------------------------------------
For iloc% = 1 To NbDeNb%
   NbEnOrdre(iloc%) = NbEnDesordre(iloc%)
Next iloc%
' ---------------------------------------------------
' ********************************
' Algorithme de tri
' ********************************
Call TriPartielSM(1, NbDeNb%, 1, NbEnOrdre())
' --------------------------------------------------
End Sub

Public Sub TriPartielAM(debpa%, finpa%, pospivpa%, Nbpa(), PermTpa%())
' ------------------------------------------------------
' Tri partiel de nbnb% nombres (avec mémoire)
' ----------------------------------------- ------------
' En entrée :  debpa%      = début de la zone à trier
'              finpa%      = fin de la zone à trier
'              Nbpa()      = tableau des nombres à trier
'              PermTra%()  =  Permutations effectuées lors du tri
'                             [PermTpi%(j%) = ancienne position du nombre
'                              placé après tri en position j%]
'
' En sortie :  pospivpa%   = position du pivot
'              Nbpa()      = tableau des nombres triés
'              PermTra%()
' ------------------------------------------------------
NbNb% = finpa% - debpa% + 1
If NbNb% < 2 Then Exit Sub
ReDim Nbloc(1 To NbNb%)
ReDim PermLoc%(1 To NbNb%)
' ------------------------------------------------------
' Algorithme de tri partiel
' ------------------------------------------------------
pivotpa = Nbpa(debpa%)
ipivpa% = 1
Nbloc(ipivpa%) = pivotpa
PermLoc%(ipivpa%) = 1
For iloc% = 2 To NbNb%
   inbpa% = debpa% + iloc% - 1
   If Nbpa(inbpa%) >= pivotpa Then
      Nbloc(iloc%) = Nbpa(inbpa%)
      PermLoc%(iloc%) = iloc%
   Else
      For jloc% = iloc% To ipivpa% + 1 Step -1
         Nbloc(jloc%) = Nbloc(jloc% - 1)
         PermLoc%(jloc%) = PermLoc%(jloc% - 1)
      Next jloc%
      Nbloc(ipivpa%) = Nbpa(inbpa%)
      PermLoc%(ipivpa%) = iloc%
      ipivpa% = ipivpa% + 1
   End If
Next iloc%
pospivpa% = debpa% + ipivpa% - 1
' ------------------------------------------------------
' Transfert de Nbloc à Nbpa et PermLoc% à PermTpa%
' ------------------------------------------------------
For iloc% = 1 To NbNb%
   Nbpa(debpa% + iloc% - 1) = Nbloc(iloc%)
   PermTpa%(debpa% + iloc% - 1) = PermLoc%(iloc%)
Next iloc%
' ----------------------------------------------------------
' Libération de la mémoire occupée par Nbloc() et PermLoc%()
' ----------------------------------------------------------
Erase Nbloc
Erase PermLoc%
' ------------------------------------------------------
' Tri partiel des 2 sous-ensembles formés
' ------------------------------------------------------
Call TriPartielAM(debpa%, pospivpa% - 1, pospivpa1%, Nbpa(), PermTpa%())
Call TriPartielAM(pospivpa% + 1, finpa%, pospivpa2%, Nbpa(), PermTpa%())
' ------------------------------------------------------
End Sub
Public Sub TriPartielSM(debpa%, finpa%, pospivpa%, Nbpa())
' ------------------------------------------------------
' Tri partiel de nbnb% nombres (sans mémoire)
' ----------------------------------------- ------------
' En entrée :  debpa%      = début de la zone à trier
'              finpa%      = fin de la zone à trier
'              Nbpa()      = tableau des nombres à trier
'
' En sortie :  pospivpa%   = position du pivot
'              Nbpa()      = tableau des nombres triés
' ------------------------------------------------------
NbNb% = finpa% - debpa% + 1
If NbNb% < 2 Then Exit Sub
ReDim Nbloc(1 To NbNb%)
' ------------------------------------------------------
' Algorithme de tri partiel
' ------------------------------------------------------
pivotpa = Nbpa(debpa%)
ipivpa% = 1
Nbloc(ipivpa%) = pivotpa
For iloc% = 2 To NbNb%
   inbpa% = debpa% + iloc% - 1
   If Nbpa(inbpa%) >= pivotpa Then
      Nbloc(iloc%) = Nbpa(inbpa%)
   Else
      For jloc% = iloc% To ipivpa% + 1 Step -1
         Nbloc(jloc%) = Nbloc(jloc% - 1)
      Next jloc%
      Nbloc(ipivpa%) = Nbpa(inbpa%)
      ipivpa% = ipivpa% + 1
   End If
Next iloc%
pospivpa% = debpa% + ipivpa% - 1
' ------------------------------------------------------
' Transfert de Nbloc à Nbpa
' ------------------------------------------------------
For iloc% = 1 To NbNb%
   Nbpa(debpa% + iloc% - 1) = Nbloc(iloc%)
Next iloc%
' ----------------------------------------------------------
' Libération de la mémoire occupée par Nbloc() et PermLoc%()
' ----------------------------------------------------------
Erase Nbloc
' ------------------------------------------------------
' Tri partiel des 2 sous-ensembles formés
' ------------------------------------------------------
Call TriPartielSM(debpa%, pospivpa% - 1, pospivpa1%, Nbpa())
Call TriPartielSM(pospivpa% + 1, finpa%, pospivpa2%, Nbpa())
' ------------------------------------------------------
End Sub

Private Sub mnuTriRapide_Click()
   ' *******************************************************
   ' affectation de leurs valeurs aux éléments de NbATrier()
   ' *******************************************************
   GridNbATrier.Col = 1
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      If GridNbATrier.Text = "" Then
         GridNbATrier.Text = "0"
         NbATrier(i%) = 0
      Else
         NbATrier(i%) = CSng(GridNbATrier.Text)
      End If
   Next i%
   ' ********************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   ' ********************************************
   Call TriRapideSM(NbNombres%, NbATrier(), NbTries())
   ' *************************************************
   ' affichage des éléments de NbTries()
   ' *************************************************
   GridNbTries.Col = 1
   For i% = 1 To NbNombres%
      GridNbTries.Row = i%
      GridNbTries.Text = Format(NbTries(i%), "0.000")
   Next i%
   ' ********************************************
   lblInfo.Caption = "Tri-rapide"
   ' ********************************************
End Sub

Private Sub txtNbNb_Change()
   On Error Resume Next
   NbNombres% = CInt(txtNbNb.Text)
   If Err.Number <> 0 Then
      If Val(txtNbNb.Text) = 0 Then
         NbNombres% = 0
      Else
         MsgBox "Nombre de nombres incorrect", 48, "POLYNOME"
         Exit Sub
      End If
   End If
   On Error GoTo 0
   If NbNombres% < 0 Then
      Beep
      MsgBox "le nombre de nombres doit être supérieur ou égal à 0 !", 48, "POLYNOME"
      NbNombres% = 0
      txtNbNb.Text = "0"
   End If
   '***************************************
   ReDim NbATrier(1 To NbNombres%)
   ReDim NbTries(1 To NbNombres%)
   '***************************************
   GridNbATrier.Rows = NbNombres% + 1
   GridNbATrier.Cols = 2
   GridNbTries.Rows = NbNombres% + 1
   GridNbTries.Cols = 2
   '*****************************************
   ' renumérotation 1ère colonne gridNbATrier
   '*****************************************
   GridNbATrier.Col = 0
   For i% = 1 To NbNombres%
      GridNbATrier.Row = i%
      GridNbATrier.Text = Format(i%, "0")
   Next i%
   '****************************************
   ' renumérotation 1ère colonne gridNbTries
   '****************************************
   GridNbTries.Col = 0
   For i% = 1 To NbNombres%
      GridNbTries.Row = i%
      GridNbTries.Text = Format(i%, "0")
   Next i%
   '*****************************************
End Sub
