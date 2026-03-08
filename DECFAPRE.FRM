VERSION 5.00
Begin VB.Form FenetreDecFaPre 
   Caption         =   "Décomposition d'un nombre entier en produit de facteurs premiers"
   ClientHeight    =   4785
   ClientLeft      =   1080
   ClientTop       =   2130
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4785
   ScaleWidth      =   8295
   Begin VB.CommandButton CommandQuitDFP 
      Caption         =   "Quitter"
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton CommandOKDec 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox TextNDec 
      Height          =   1335
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   6255
   End
   Begin VB.TextBox TextNombre 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label LabelNDec 
      Caption         =   "N ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label LabelN 
      Caption         =   "N :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label LabelNombre 
      Caption         =   "Nombre "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FenetreDecFaPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandOKDec_Click()
   Call DecFaPre
End Sub

Private Sub CommandQuitDFP_Click()
   FenetreDecFaPre.Hide
End Sub

Private Sub DecFaPre()
   On Error GoTo Traite_ErreursDPFP
   TextNDec.Text = "Calcul en cours..."
   NDec$ = ""
   Ndfp# = CDbl(TextNombre.Text)
   i# = 0
   Mdfp# = Ndfp#
   Pdfp# = 1
   Ldfp# = Sqr(Mdfp#)
   Do
   ex# = 0
   Pdfp# = Pdfp# + 1
   If Pdfp# > 3 Then Pdfp# = Pdfp# + 1
   Do
      Qdfp# = Mdfp# / Pdfp#
      If Qdfp# > Int(Qdfp#) Then
         If ex# > 0 Then
            If i# > 0 Then NDec$ = NDec$ + " x "
            NDec$ = NDec$ + Format(Pdfp#)
            i# = i# + 1
            If ex# > 1 Then NDec$ = NDec$ + "^" + Format(ex#)
         End If
         Exit Do
      End If
      ex# = ex# + 1
      Mdfp# = Qdfp#
   Loop
   Loop While Pdfp# < Ldfp#
   If Mdfp# > 1 Then
   If i# > 0 Then NDec$ = NDec$ + " x "
   NDec$ = NDec$ + Format(Mdfp#)
   End If
   TextNDec.Text = NDec$
   Exit Sub
Traite_ErreursDPFP:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub TextNombre_Change()
    TextNDec.Text = ""
End Sub
