VERSION 5.00
Begin VB.Form FenetreDerive 
   Caption         =   "Dérivée d'une expression numérique"
   ClientHeight    =   3915
   ClientLeft      =   690
   ClientTop       =   1770
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
   Begin VB.TextBox txtVar 
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
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtExpression 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label lblExpressionDerivee 
      Caption         =   "Expression dérivée :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblExprDerivee 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Label lbldF 
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
      Width           =   975
   End
   Begin VB.Label lblVariable 
      Caption         =   "Variable de dérivation :"
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
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblF 
      Caption         =   "F ="
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
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblExpr 
      Caption         =   "Expression à dériver :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreDerive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Private Sub Form_Activate()
   txtVar.Text = nomvar$
   txtExpression.Text = expression$
   lblExprDerivee.Caption = ""
End Sub

Private Sub Form_Deactivate()
   nomvar$ = txtVar.Text
   expression$ = txtExpression.Text
   exprderivee = lblExprDerivee.Caption
End Sub

Private Sub Form_Load()
   '-----------------------------------------------------
   '---------------- valeurs par défaut -----------------
   '----------------    pour Dérivée    -----------------
   '-----------------------------------------------------
   nomvar$ = "X"
   expression$ = "X^X"
   txtVar.Text = nomvar$
   txtExpression.Text = expression$
   lbldF.Caption = "dF/dX ="
   lblExprDerivee.Caption = ""
   ' solution exacte : dF/dX = X^X*LOG(X)+X^(X-1)
End Sub







Private Sub mnuCalculer_Click()
   If lblExprDerivee.Caption = "" Then
      '--------------------------------------------------------------
      '--------- Première manipulation de l'expression :    ---------
      '--------- Suppression des blancs dans les formules,  ---------
      '--------- vérification du nombre de parenthèses et   ---------
      '--------------------------------------------------------------
      Call OteBlancs(UCase(expression$), expression$)
      If Erreur = True Then Exit Sub
      '--------------------------------------------------------------
      '--------- Dérivation    ---------
      '--------------------------------------------------------------
      Call DeriveSomme(expression$, nomvar$, exprderivee$)
      If Erreur = True Then Exit Sub
      '--------------------------------------------------------------
      '--------- Affichage du résultat ---------
      '--------------------------------------------------------------
      lblExprDerivee.Caption = exprderivee$
   End If
End Sub



Private Sub mnuQuitter_Click()
   FenetreDerive.Hide
End Sub





Private Sub txtExpression_Change()
   expression$ = txtExpression.Text
   lblExprDerivee.Caption = ""
End Sub



Private Sub txtVar_Change()
   On Error Resume Next
   nomvar$ = CStr(txtVar.Text)
   lbldF.Caption = "dF/d" & nomvar$ & " = "
   lblExprDerivee.Caption = ""
   On Error GoTo 0
End Sub


