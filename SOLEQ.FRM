VERSION 5.00
Begin VB.Form FenetreSolEq 
   Caption         =   "Solution de l'équation F(X)=0"
   ClientHeight    =   5085
   ClientLeft      =   720
   ClientTop       =   1590
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
   ScaleHeight     =   5085
   ScaleWidth      =   7515
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
      TabIndex        =   0
      Text            =   "0"
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
      TabIndex        =   1
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lblMessage 
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
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label lblMethode 
      Caption         =   " Recherche des solutions de l'équation F(X)=0 par la méthode de NEWTON-RAPHSON."
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
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblSolEq 
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
      TabIndex        =   2
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label lblI 
      Caption         =   "X ="
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
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblXa 
      Caption         =   "Xa ="
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
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblValAp 
      Caption         =   "Valeur approchée de la solution :"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label EtiquetteVar2 
      Caption         =   "F(X) ="
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
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblFonction 
      Caption         =   "Fonction dont on cherche la racine :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   3255
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer"
   End
   Begin VB.Menu mnuTracer 
      Caption         =   "&Tracer"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreSolEq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Private Sub DefVarf_Change()
   FenetreSolEq.lblMessage.Caption = ""
End Sub

Private Sub Form_Activate()
   FenetreSolEq.txtXa.Text = Format(Xapproche, "0.000000")
   FenetreSolEq.DefVarf.Text = TextVar2$
   FenetreSolEq.lblSolEq.Caption = ""
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Xapproche = CSng(txtXa.Text)
   If Err.Number <> 0 Then
      MsgBox "Xa est incorrect", 48, "SOLEQ"
      Exit Sub
   End If
   On Error GoTo 0
   TextVar2$ = DefVarf.Text
End Sub

Private Sub Form_Load()
   Xapproche = 4
   txtXa.Text = Format(Xapproche, "0.000000")
   TextVar2$ = "0,2*X^5-0,6*X^4-X^3+3*X^2+0,8*X-2,4"
   DefVarf.Text = TextVar2$
   'Xapproche = 4
   ' La solution de F(X) = 0
   ' avec F(X)=0,2*X^5-0,6*X^4-X^3+3*X^2+0,8*X-2,4
   ' et Xapproche = 4
   ' est : Xsolution = 3
End Sub


Private Sub mnuCalculer_Click()
   On Error Resume Next
   Xapproche = CSng(txtXa.Text)
   If Err.Number <> 0 Then
      MsgBox "Xa est incorrect", 48, "SOLEQ"
      Exit Sub
   End If
   On Error GoTo 0
   '--------------------------------------------------------------
   FenetreSolEq.lblMessage.ForeColor = ROUGE
   FenetreSolEq.lblMessage.Caption = " CALCUL EN COURS ..."
   DoEvents
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   NombVar% = 1
   var$(1) = "X"
   TextVar2$ = FenetreSolEq.DefVarf.Text
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   If Erreur = True Then
      MsgBox "Erreur dans la formule !", 48, "SolEq"
      Exit Sub
   End If
   '--------------------------------------------------------------
   Call NewtonRaphson
   '--------------------------------------------------------------
   FenetreSolEq.lblMessage.ForeColor = MAGENTA
   FenetreSolEq.lblMessage.Caption = Message$
   FenetreSolEq.lblSolEq.ForeColor = ROUGE
   FenetreSolEq.lblSolEq.Caption = Format(Xsolution, "0.00000")
End Sub


Private Sub mnuQuitter_Click()
   FenetreSolEq.Hide
End Sub


Private Sub mnuTracer_Click()
   On Error Resume Next
   Xapproche = CSng(txtXa.Text)
   If Err.Number <> 0 Then
      MsgBox "Xa est incorrect", 48, "SOLEQ"
      Exit Sub
   End If
   On Error GoTo 0
   If FenetreSolEq.lblSolEq.Caption = "" Then
      mnuCalculer_Click
   End If
   FenetreSolEq.Hide
   FenetreDefSolEq.Show
End Sub
Private Sub txtXa_Change()
   On Error Resume Next
   Xapproche = CSng(txtXa.Text)
   FenetreSolEq.lblMessage.Caption = ""
   On Error GoTo 0
End Sub



