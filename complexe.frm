VERSION 5.00
Begin VB.Form FenetreComplexe 
   Caption         =   "Nombres Complexes"
   ClientHeight    =   6105
   ClientLeft      =   315
   ClientTop       =   720
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   8985
   Begin VB.TextBox txtArgZ2 
      Height          =   375
      Left            =   6720
      TabIndex        =   28
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtModZ2 
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtImZ2 
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtReZ2 
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtArgZ1 
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtModZ1 
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtImZ1 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtQuotient 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   5535
   End
   Begin VB.TextBox txtProduit 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox txtReZ1 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblEgale2 
      Caption         =   "= "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblI2 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   34
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblI1 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   33
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblIPi2 
      Caption         =   "P)"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   32
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblIf2 
      Alignment       =   1  'Right Justify
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   31
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblIf1 
      Alignment       =   1  'Right Justify
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   30
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblArgZ2 
      Caption         =   "argument :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblPlusExpI2 
      Caption         =   "exp("
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblModZ2 
      Caption         =   "module :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   26
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblImZ2 
      Caption         =   "partie imaginaire :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblPlus2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblArgZ1 
      Caption         =   "argument :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   21
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblModZ1 
      Caption         =   "module :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblIPi1 
      Caption         =   "P)"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblPlusExpI1 
      Caption         =   "exp("
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblEgale1 
      Caption         =   "= "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblImZ1 
      Caption         =   "partie imaginaire :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblPlus1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblReZ2 
      Caption         =   "partie réelle :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LabelZ1surZ2 
      Caption         =   "z1 / z2 ="
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
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabelZ1Z2 
      Caption         =   "z1 . z2 ="
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
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LabelQuotient 
      Caption         =   "Quotient :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabelProduit 
      Caption         =   "Produit :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblReZ1 
      Caption         =   "partie réelle :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblZ2 
      Caption         =   "z2 ="
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblZ1 
      Caption         =   "z1 ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreComplexe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReZ1, ImZ1, ModZ1, ArgZ1
Dim ReZ2, ImZ2, ModZ2, ArgZ2
Dim ReZp, ImZp, ModZp, ArgZp
Dim ReZq, ImZq, ModZq, ArgZq
Dim EviteReactionEnChaine As Boolean
Dim DivisionParZero As Boolean


Public Sub DivisionComplexe(PartR1, PartI1, PartR2, PartI2, PartRq, PartIq)
   ' *************************************************************************
   ' Division complexe : PartRq+i.PartIq = (PartR1+i.PartI1)/(PartR2+i.PartI2)
   ' *************************************************************************
   DivisionParZero = False
   ModCar2 = PartR2 * PartR2 + PartI2 * PartI2
   If ModCar2 = 0 Then
      DivisionParZero = True
   Else
      PartRq = (PartR1 * PartR2 + PartI1 * PartI2) / ModCar2
      PartIq = (PartI1 * PartR2 - PartR1 * PartI2) / ModCar2
   End If
End Sub

Public Sub ProduitComplexe(PartR1, PartI1, PartR2, PartI2, PartRp, PartIp)
   ' Produit complexe : PartRp+i.PartIp = (PartR1+i.PartI1)*(PartR2+i.PartI2)
   PartRp = PartR1 * PartR2 - PartI1 * PartI2
   PartIp = PartR1 * PartI2 + PartI1 * PartR2


' CALCULS SUR LES NOMBRES COMPLEXES
' Initialisation :
'eps = 0.000001
'PI = 3.1415926536
'Screen 0, 0
'Cls
'Print
'Print "                 CALCULS SUR LES NOMBRES COMPLEXES"
'Print
'Print
'Print "                                                           i . tˆta"
'Print "z = x + i . y = rh“ . [cos(tˆta) + i . sin(tˆta)] = rh“ . e"
'Print
'Print "                                       0.5        "
'Print "On a :  module de z   = |z| = (xý + yý)    = rh“  "
'Print
'Print "        argument de z = Arg(z) = tˆta"
'Print
'Print
'Do
   ' *****************************************************
   ' effacement des lignes 14 … 25
   ' *****************************************************
'   LOCATE 13
'   For ligne% = 14 To 25
'      Print Space$(80);
'   Next ligne%
  ' *****************************************************
   ' Menu principal
   ' *****************************************************
'   LOCATE 14, 1
'   Print "Voulez-vous :"
'   Print
'   Print " 0) arrˆter"
'   Print " 1) passer de (x,y) … (rh“, tˆta)"
'  Print " 2) passer de (rh“, tˆta) … (x,y)"
'   Print " 3) multiplier deux complexes"
'   Print " 4) diviser deux complexes"
'   Do
'      a$ = INKEY$
'   Loop While a$ = ""
'   ' *****************************************************
   ' effacement des lignes 14 … 25
   ' *****************************************************
'   LOCATE 13
'   For ligne% = 14 To 25
'     Print Space$(80);
'   Next ligne%
'   ' *****************************************************
   ' action en fonction du choix
   ' *****************************************************
'   choix% = Val(a$)
'   Select Case choix%
'   Case 0
'      Cls
'      End
'   Case 1
'      ' *****************************************************
'      ' passer de (x,y) … (rh“, tˆta)"
'      ' *****************************************************
'      LOCATE 14, 1
'      INPUT "x = ", x
'      INPUT "y = ", y
'      ' *****************************************************
      ' calcul de rh“ et tˆta … partir de x et y
      ' *****************************************************
'      rho = Sqr(X * X + Y * Y)
'      If Abs(Y) < eps Then
'         If Abs(X) < eps Then
'            TETA = 0
'         ElseIf X < 0 Then
'            TETA = -PI
'         Else
'            TETA = 0
'         End If
'      ElseIf Y > 0 Then
'         If Abs(X) < eps Then
'            TETA = PI / 2
'         ElseIf X > 0 Then
'            TETA = Atn(X / Y)
'         Else
'            TETA = PI + Atn(X / Y)
'         End If
'      Else
'         If Abs(X) < eps Then
'            TETA = 3 * PI / 2
'         ElseIf X > 0 Then
'            TETA = 2 * PI + Atn(X / Y)
'         Else
'            TETA = PI + Atn(X / Y)
'         End If
'      End If
      ' *****************************************************
      ' pr‚sentation des r‚sultats
      ' *****************************************************
'      Print
'      Print "rh“  = "; rho
'      Print "tˆta = "; TETA
'      ' *****************************************************
'      ' suite
'      ' *****************************************************
'      Print
'      Print "frappez une touche pour continuer"
'      Do
'         k$ = INKEY$
'      Loop While k$ = ""
'      ' *****************************************************
'   Case 2
      ' *****************************************************
      ' passer de (x,y) … (rh“, tˆta)"
      ' *****************************************************
'      LOCATE 14, 1
'      INPUT "rh“  = ", rho
'      INPUT "tˆta = ", teta
      ' *****************************************************
      ' calcul de x et y … partir de rh“ et tˆta
      ' *****************************************************
'      X = rho * Cos(TETA)
'      Y = rho * Sin(TETA)
      ' *****************************************************
      ' pr‚sentation des r‚sultats
      ' *****************************************************
'      Print
'      Print "x = "; X
'      Print "y = "; Y
      ' *****************************************************
      ' suite
      ' *****************************************************
'      Print
'      Print "frappez une touche pour continuer"
'      Do
'         k$ = INKEY$
'      Loop While k$ = ""
'      ' *****************************************************
'   Case 3
      ' *****************************************************
      ' multiplier deux complexes
      ' *****************************************************
'      LOCATE 14, 1
'      Print "Premier complexe z1 :"
'      INPUT "x1 = ", x1
'      INPUT "y1 = ", y1
'      Print "DeuxiŠme complexe z2 :"
'      INPUT "x2 = ", x2
'      INPUT "y2 = ", y2
      ' *****************************************************
      ' calcul de z1.z2 … partir de (x1,x2) et (y1,y2)
      ' *****************************************************
'      X = X1 * X2 - Y1 * Y2
'      Y = X1 * Y2 + X2 * Y1
      ' *****************************************************
      ' pr‚sentation des r‚sultats
      ' *****************************************************
'      Print
'      Print "z = z1 . z2 =  "; X; " + "; Y; " . i"
      ' *****************************************************
      ' suite
      ' *****************************************************
'      Print
'      Print "frappez une touche pour continuer"
'      Do
'         k$ = INKEY$
'      Loop While k$ = ""
      ' *****************************************************
'   Case 4
      ' *****************************************************
      ' diviser deux complexes
      ' *****************************************************
'      LOCATE 14, 1
'      Print "Premier complexe z1 :"
'      INPUT "x1 = ", x1
'      INPUT "y1 = ", y1
'      Print "DeuxiŠme complexe z2 :"
'      INPUT "x2 = ", x2
'      INPUT "y2 = ", y2
      ' *****************************************************
      ' calcul de z1/z2 … partir de (x1,x2) et (y1,y2)
      ' *****************************************************
'      If Abs(X2) < eps And Abs(Y2) < eps Then
'         Print
'         Print "z = l'infini"
'      Else
'         den = X2 * X2 + Y2 * Y2
'         X = (X1 * X2 + Y1 * Y2) / den
'         Y = (X2 * Y1 - X1 * Y2) / den
   ' *****************************************************
   ' pr‚sentation des r‚sultats
   ' *****************************************************
'         Print
'         Print "z = z1 / z2 =  "; X; " + "; Y; " . i"
'      End If
      ' *****************************************************
      ' suite
      ' *****************************************************
'      Print
'      Print "frappez une touche pour continuer"
'      Do
'         k$ = INKEY$
'      Loop While k$ = ""
      ' *****************************************************
'   Case Else
'      Beep
'   End Select
'Loop
End Sub

Public Sub ReImaModArg(ReZ, ImZ, ModZ, ArgZ)
   ' *********************************************
   ' Calcule le module et l'argument d'un complexe
   ' à partir de ses parties réelle et imaginaire
   ' *********************************************
   eps = 0.000001
   ModZ = Sqr(ReZ * ReZ + ImZ * ImZ)
   If Abs(ImZ) < eps Then
      If Abs(ReZ) < eps Then
         ArgZ = 0
      ElseIf ReZ < 0 Then
         ArgZ = -PI
      Else
         ArgZ = 0
      End If
   ElseIf ImZ > 0 Then
      If Abs(ReZ) < eps Then
         ArgZ = PI / 2
      ElseIf ReZ > 0 Then
         ArgZ = Atn(ReZ / ImZ)
      Else
         ArgZ = PI + Atn(ReZ / ImZ)
      End If
   Else
      If Abs(ReZ) < eps Then
         ArgZ = 3 * PI / 2
      ElseIf ReZ > 0 Then
         ArgZ = 2 * PI + Atn(ReZ / ImZ)
      Else
         ArgZ = PI + Atn(ReZ / ImZ)
      End If
   End If
End Sub

Public Sub ModArgaReIm(ReZ, ImZ, ModZ, ArgZ)
   ' ******************************************************
   ' Calcule les parties réelle et imaginaire d'un complexe
   ' à partir de son module et de son argument
   ' ******************************************************
   ReZ = ModZ * Cos(ArgZ)
   ImZ = ModZ * Sin(ArgZ)
End Sub

Private Sub Form_Load()
   ' Mise en place des valeurs par défaut
   ReZ1 = 1
   ImZ1 = 2
   txtReZ1.Text = Format(ReZ1, "0.0000")
   txtImZ1.Text = Format(ImZ1, "0.0000")
   ReZ2 = 3
   ImZ2 = 4
   txtReZ2.Text = Format(ReZ2, "0.0000")
   txtImZ2.Text = Format(ImZ2, "0.0000")
   'txtProduit.Text = "-5,0000 + 10,0000 * i = 11,1803 * exp(2,6779 * i * PI)"
   'txtQuotient.Text = "0,4400 + 0,0800 * i = 0,4472 * exp(1,3909 * i * PI)"
End Sub

Private Sub mnuQuitter_Click()
   FenetreComplexe.Hide
End Sub


Private Sub txtArgZ1_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ArgZ1 = CSng(txtArgZ1.Text)
   ArgZ1 = ArgZ1 * PI
   On Error GoTo 0
   Call ModArgaReIm(ReZ1, ImZ1, ModZ1, ArgZ1)
   txtReZ1.Text = Format(ReZ1, "0.0000")
   txtImZ1.Text = Format(ImZ1, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub

Private Sub txtArgZ2_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ArgZ2 = CSng(txtArgZ2.Text)
   ArgZ2 = ArgZ2 * PI
   On Error GoTo 0
   Call ModArgaReIm(ReZ2, ImZ2, ModZ2, ArgZ2)
   txtReZ2.Text = Format(ReZ2, "0.0000")
   txtImZ2.Text = Format(ImZ2, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub


Private Sub txtImZ1_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ImZ1 = CSng(txtImZ1.Text)
   On Error GoTo 0
   Call ReImaModArg(ReZ1, ImZ1, ModZ1, ArgZ1)
   txtModZ1.Text = Format(ModZ1, "0.0000")
   txtArgZ1.Text = Format(ArgZ1 / PI, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub

Private Sub txtImZ2_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ImZ2 = CSng(txtImZ2.Text)
   On Error GoTo 0
   Call ReImaModArg(ReZ2, ImZ2, ModZ2, ArgZ2)
   txtModZ2.Text = Format(ModZ2, "0.0000")
   txtArgZ2.Text = Format(ArgZ2 / PI, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub


Private Sub txtModZ1_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ModZ1 = CSng(txtModZ1.Text)
   On Error GoTo 0
   Call ModArgaReIm(ReZ1, ImZ1, ModZ1, ArgZ1)
   txtReZ1.Text = Format(ReZ1, "0.0000")
   txtImZ1.Text = Format(ImZ1, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub


Private Sub txtModZ2_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ModZ2 = CSng(txtModZ2.Text)
   On Error GoTo 0
   Call ModArgaReIm(ReZ2, ImZ2, ModZ2, ArgZ2)
   txtReZ2.Text = Format(ReZ2, "0.0000")
   txtImZ2.Text = Format(ImZ2, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub


Private Sub txtReZ1_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ReZ1 = CSng(txtReZ1.Text)
   On Error GoTo 0
   Call ReImaModArg(ReZ1, ImZ1, ModZ1, ArgZ1)
   txtModZ1.Text = Format(ModZ1, "0.0000")
   txtArgZ1.Text = Format(ArgZ1 / PI, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub


Private Sub txtReZ2_Change()
If EviteReactionEnChaine = False Then
   EviteReactionEnChaine = True
   On Error Resume Next
   ReZ2 = CSng(txtReZ2.Text)
   On Error GoTo 0
   Call ReImaModArg(ReZ2, ImZ2, ModZ2, ArgZ2)
   txtModZ2.Text = Format(ModZ2, "0.0000")
   txtArgZ2.Text = Format(ArgZ2 / PI, "0.0000")
   EviteReactionEnChaine = False
End If
Call ProduitComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZp, ImZp)
Call ReImaModArg(ReZp, ImZp, ModZp, ArgZp)
ZProd$ = Format(ReZp, "0.0000") & " + " & Format(ImZp, "0.0000") & " * i"
ZProd$ = ZProd$ & " = " & Format(ModZp, "0.0000") & " * exp("
ZProd$ = ZProd$ & Format(ArgZp, "0.0000") & " * i * PI)"
txtProduit.Text = ZProd$
Call DivisionComplexe(ReZ1, ImZ1, ReZ2, ImZ2, ReZq, ImZq)
If DivisionParZero = True Then
   txtQuotient.Text = "Division par zéro !"
Else
   Call ReImaModArg(ReZq, ImZq, ModZq, ArgZq)
   ZQuot$ = Format(ReZq, "0.0000") & " + " & Format(ImZq, "0.0000") & " * i"
   ZQuot$ = ZQuot$ & " = " & Format(ModZq, "0.0000") & " * exp("
   ZQuot$ = ZQuot$ & Format(ArgZq, "0.0000") & " * i * PI)"
   txtQuotient.Text = ZQuot$
End If
End Sub


