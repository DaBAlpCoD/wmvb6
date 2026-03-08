VERSION 5.00
Begin VB.Form FenetreIntegrale 
   Caption         =   "Intégrale d'une fonction"
   ClientHeight    =   3915
   ClientLeft      =   675
   ClientTop       =   1125
   ClientWidth     =   7515
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
   ScaleHeight     =   3915
   ScaleWidth      =   7515
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
      TabIndex        =   0
      Top             =   1800
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
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
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   6615
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
      Left            =   480
      TabIndex        =   3
      Top             =   3240
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
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   495
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
      Left            =   240
      TabIndex        =   5
      Top             =   1800
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
      Left            =   240
      TabIndex        =   6
      Top             =   1560
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
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
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
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   4695
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
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblFonction 
      Caption         =   "Fonction à intégrer :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   2055
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
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
Attribute VB_Name = "FenetreIntegrale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Private Sub DefVarf_Change()
   TextVar2$ = DefVarf.Text
   FenetreDefInteg.DefVarf.Text = TextVar2$
   lblInteg.Caption = ""
End Sub

Private Sub Form_Activate()
   txtXa.Text = Format(Xa, "0.000000")
   txtXb.Text = Format(Xb, "0.000000")
   DefVarf.Text = TextVar2$
   lblInteg.Caption = ""
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Xa = CSng(txtXa.Text)
   If Err.Number <> 0 Then
      MsgBox "Xa est incorrect", 48, "Intégrale"
      Exit Sub
   End If
   Xb = CSng(txtXb.Text)
   If Err.Number <> 0 Then
      MsgBox "Xb est incorrect", 48, "Intégrale"
      Exit Sub
   End If
   lblSolEq = ""
   TextVar2$ = DefVarf.Text
   On Error GoTo 0
End Sub

Private Sub Form_Load()
   '-----------------------------------------------------
   '---------------- valeurs par défaut -----------------
   '----------------   pour Integrale   -----------------
   '-----------------------------------------------------
   Xa = 0
   Xb = 1
   lblInteg.Caption = ""
   TextVar2$ = "X"
   DefVarf.Text = TextVar2$
   ' solution exacte : I=1/2
End Sub



Private Sub mnuCalculer_Click()
   If lblInteg.Caption = "" Then
      '-----------------------------------------------------
      '---------   Vérifications avant calcul---------------
      '-----------------------------------------------------
      On Error Resume Next
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
      '-----------------------------------------------------
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
      If Erreur = True Then Exit Sub
      '--------------------------------------------------------------
      On Error GoTo Traite_ErreursInteg
      Call Simpson
      lblInteg.ForeColor = NOIR
      lblInteg.Caption = Format(Integrale, "0.0000")
      Exit Sub
Traite_ErreursInteg:
      Select Case Err
         ' erreur 5  : appel de fonction illégal (illegal function call)
         ' erreur 6  : nombre trop grand (overflow)
         ' erreur 11 : division par zéro
         Case 5
            MsgBox "Fonction non définie pour cette valeur !", 48, "INTEGRALE"
            lblInteg.Caption = ""
         Case 6
            MsgBox "Nombre trop grand !", 48, "INTEGRALE"
            lblInteg.Caption = ""
         Case 11
            MsgBox "Division par zéro !", 48, "INTEGRALE"
            lblInteg.Caption = ""
         Case Else
            Resume
      End Select
      Exit Sub
   End If
End Sub


Private Sub mnuQuitter_Click()
   FenetreIntegrale.Hide
End Sub


Private Sub mnuTracer_Click()
   FenetreIntegrale.Hide
   FenetreDefInteg.Show
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



