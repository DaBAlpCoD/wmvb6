VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreProLin 
   Caption         =   "Programmation linéaire"
   ClientHeight    =   5985
   ClientLeft      =   195
   ClientTop       =   1005
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
   ScaleHeight     =   5985
   ScaleWidth      =   9045
   Begin MSFlexGridLib.MSFlexGrid GridEpl 
      Height          =   1335
      Left            =   1920
      TabIndex        =   33
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid GridXpl 
      Height          =   1335
      Left            =   120
      TabIndex        =   32
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid GridCoutsMar 
      Height          =   1335
      Left            =   3840
      TabIndex        =   31
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid GridCpl 
      Height          =   1335
      Left            =   6480
      TabIndex        =   30
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid GridBpl 
      Height          =   1335
      Left            =   4440
      TabIndex        =   29
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid GridApl 
      Height          =   1215
      Left            =   120
      TabIndex        =   28
      Top             =   2520
      Width           =   3855
      _ExtentX        =   6800
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
   Begin VB.TextBox txtNbCon 
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
      Left            =   6360
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtNbVar 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblBornes 
      Caption         =   "Bornes :"
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblCoutsMar 
      Caption         =   "Coûts marginaux :"
      Height          =   495
      Left            =   3840
      TabIndex        =   25
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblNbIter 
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
      Left            =   7920
      TabIndex        =   24
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblNombIter 
      Caption         =   "Nombre d'itérations :"
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
      Left            =   5760
      TabIndex        =   23
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lblValF 
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
      Left            =   7920
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
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
      Height          =   255
      Left            =   7440
      TabIndex        =   21
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblFO 
      Caption         =   "Valeur de la fonction objectif :"
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label lblExplic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label lblE 
      Caption         =   "E(i)"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblVarEc 
      Caption         =   "Variables d'écart :"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblX 
      Caption         =   "X(j)"
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
      Left            =   720
      TabIndex        =   11
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblVar 
      Caption         =   "Variables :"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblC 
      Caption         =   "C(j)"
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
      Left            =   8160
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblB 
      Caption         =   "B(i)"
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
      Left            =   5400
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblA 
      Caption         =   "A(i,j)"
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
      Left            =   2880
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblCoefFObj 
      Caption         =   "Coefficients de la fonction objectif :"
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblNbCon 
      Caption         =   "Nombre de contraintes :"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblNbVar 
      Caption         =   "Nombre de variables :"
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
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblCoefCon 
      Caption         =   "Coefficients des contraintes :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
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
Attribute VB_Name = "FenetreProLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xpl()
Dim Cpl()
Dim Apl()
Dim ANpl()
Dim Bpl()
Dim Epl()
Dim NLpl%, NCpl%
Dim NbCon%, NbVar%




Private Sub Form_Load()
   '***************************************
   '              Explications
   '***************************************
   PLText$ = "On cherche à rendre maximale la fonction objectif "
   PLText$ = PLText$ & "F=C(1) X(1)+...+C(N) X(N)" & Chr$(13)
   PLText$ = PLText$ & "sachant que les N variables sont positives "
   PLText$ = PLText$ & "et soumises aux M contraintes :" & Chr$(13)
   PLText$ = PLText$ & "A(I,1)*X(1)+...+A(I,N)*X(N) <= B(I)  "
   PLText$ = PLText$ & "(avec I=1,2,...,M)" & Chr$(13)
   PLText$ = PLText$ & "On introduit pour cela M variables d'écart E(I) "
   PLText$ = PLText$ & "telles que :" & Chr$(13)
   PLText$ = PLText$ & "A(I,1)*X(1)+...+A(I,N)*X(N)+E(I) = B(I) "
   PLText$ = PLText$ & "(avec I=1,2,...,M)" & Chr$(13) & Chr$(13)
   PLText$ = PLText$ & "La résolution se fait en utilisant l'algorithme "
   PLText$ = PLText$ & "du SIMPLEXE, par la méthode des tableaux."
   lblExplic.Caption = PLText$
   '***************************************
   NbVar% = 2
   NbCon% = 3
   NLpl% = NbCon% + 1
   NCpl% = NbVar% + NbCon% + 1
   '***************************************
   ReDim Apl(1 To NLpl%, 1 To NCpl%)
   ReDim ANpl(1 To NLpl%, 1 To NCpl%)
   ReDim Bpl(1 To NbCon%)
   ReDim Cpl(1 To NbVar%)
   ReDim Epl(1 To NbCon%)
   ReDim Xpl(1 To NbVar%)
   '***************************************
   txtNbVar.Text = Format(NbVar%, "0")
   txtNbCon.Text = Format(NbCon%, "0")
   '***************************************
   '***** mise en place des 5 grilles *****
   '***************************************
   gridApl.Rows = NbCon% + 1
   gridApl.Cols = NbVar% + 1
   gridBpl.Rows = NbCon% + 1
   gridBpl.Cols = 2
   gridCpl.Rows = NbVar% + 1
   gridCpl.Cols = 2
   gridEpl.Rows = NbCon% + 1
   gridEpl.Cols = 2
   gridCoutsMar.Rows = NbCon% + 1
   gridCoutsMar.Cols = 2
   gridXpl.Rows = NbVar% + 1
   gridXpl.Cols = 2
   '***************************************
   ' alignements et largeurs colonnes
   '***************************************
   For i% = 0 To NbVar%
      FenetreProLin.gridApl.FixedAlignment(i%) = 2
      FenetreProLin.gridApl.ColWidth(i%) = 800
   Next i%
   For i% = 0 To 1
      FenetreProLin.gridBpl.FixedAlignment(i%) = 2
      FenetreProLin.gridBpl.ColWidth(i%) = 800
      FenetreProLin.gridCpl.FixedAlignment(i%) = 2
      FenetreProLin.gridCpl.ColWidth(i%) = 800
      FenetreProLin.gridEpl.FixedAlignment(i%) = 2
      FenetreProLin.gridEpl.ColWidth(i%) = 800
      FenetreProLin.gridCoutsMar.FixedAlignment(i%) = 2
      FenetreProLin.gridCoutsMar.ColWidth(i%) = 800
      FenetreProLin.gridXpl.FixedAlignment(i%) = 2
      FenetreProLin.gridXpl.ColWidth(i%) = 800
   Next i%
   ' ********************************
   ' numérotation 1ères lignes
   ' ********************************
   gridApl.Row = 0
   For i% = 1 To NbVar%
      gridApl.Col = i%
      gridApl.Text = i%
   Next i%
   '**********************************
   gridBpl.Row = 0
   gridBpl.Col = 1
   gridBpl.Text = "B"
   '*******************
   gridCpl.Row = 0
   gridCpl.Col = 1
   gridCpl.Text = "C"
   '*******************
   gridEpl.Row = 0
   gridEpl.Col = 1
   gridEpl.Text = "E"
   '*******************
   gridCoutsMar.Row = 0
   gridCoutsMar.Col = 1
   gridCoutsMar.Text = ""
   '*******************
   gridXpl.Row = 0
   gridXpl.Col = 1
   gridXpl.Text = "X"
   ' *********************************
   ' numérotation 1ères colonnes
   ' *********************************
   gridApl.Col = 0
   gridBpl.Col = 0
   gridCpl.Col = 0
   gridEpl.Col = 0
   gridCoutsMar.Col = 0
   gridXpl.Col = 0
   For i% = 1 To NbVar%
      gridApl.Row = i%
      gridApl.Text = i%
      gridCpl.Row = i%
      gridCpl.Text = i%
      gridXpl.Row = i%
      gridXpl.Text = i%
   Next i%
   For i% = 1 To NbCon%
      gridBpl.Row = i%
      gridBpl.Text = i%
      gridEpl.Row = i%
      gridEpl.Text = i%
      gridCoutsMar.Row = i%
      gridCoutsMar.Text = i%
   Next i%
   ' ****************************************
   '       Valeurs par défaut
   ' ****************************************
   Apl(1, 1) = 2
   Apl(1, 2) = 3
   Bpl(1) = 18
   Apl(2, 1) = 2
   Apl(2, 2) = 1
   Bpl(2) = 10
   Apl(3, 1) = 1
   Apl(3, 2) = 0
   Bpl(3) = 4
   '***************************************
   Cpl(1) = 4
   Cpl(2) = 3
   ' ****************************************
   ' placement des valeurs par défaut de Apl
   ' ****************************************
   For i% = 1 To NbCon%
      For j% = 1 To NbVar%
         FenetreProLin.gridApl.Row = i%
         FenetreProLin.gridApl.Col = j%
         FenetreProLin.gridApl.Text = Format(Apl(i%, j%), "0.000")
      Next j%
   Next i%
   ' ****************************************
   ' placement des valeurs par défaut de Bpl
   ' ****************************************
   FenetreProLin.gridBpl.Col = 1
   For i% = 1 To NbCon%
      FenetreProLin.gridBpl.Row = i%
      FenetreProLin.gridBpl.Text = Format(Bpl(i%), "0.000")
   Next i%
   ' ****************************************
   ' placement des valeurs par défaut de Cpl
   ' ****************************************
   FenetreProLin.gridCpl.Col = 1
   For i% = 1 To NbVar%
      FenetreProLin.gridCpl.Row = i%
      FenetreProLin.gridCpl.Text = Format(Cpl(i%), "0.000")
   Next i%
   ' ****************************************
End Sub

Private Sub gridApl_KeyPress(KeyAscii As Integer)
   EleMText$ = gridApl.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      gridApl.Text = EleMText$
   Case 8
      If Len(gridApl.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         gridApl.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub gridBpl_KeyPress(KeyAscii As Integer)
   EleMText$ = gridBpl.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      gridBpl.Text = EleMText$
   Case 8
      If Len(gridCpl.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         gridBpl.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub gridCpl_KeyPress(KeyAscii As Integer)
   EleMText$ = gridCpl.Text
   Select Case KeyAscii
   Case 32 To 168
      EleMCar$ = Chr(KeyAscii)
      EleMText$ = EleMText$ & EleMCar$
      gridCpl.Text = EleMText$
   Case 8
      If Len(gridCpl.Text) > 0 Then
         EleMText$ = Left$(EleMText$, Len(EleMText$) - 1)
         gridCpl.Text = EleMText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub Simplexe()
   ' ********************************************
   '               ReDims
   ' ********************************************
   Erase Apl, Bpl, Cpl
   ReDim Apl(1 To NLpl%, 1 To NCpl%)
   ReDim ANpl(1 To NLpl%, 1 To NCpl%)
   ReDim Bpl(1 To NbCon%)
   ReDim Cpl(1 To NbVar%)
   ReDim Epl(1 To NbCon%)
   ReDim Xpl(1 To NbVar%)
   ' *************************************************
   '               Message
   ' *************************************************
   lblInfo.Caption = "CALCUL EN COURS..."
   DoEvents
   ' *************************************************
   '               Initialisations
   ' *************************************************
   For i% = 1 To NbVar%
      Xpl(i%) = 0
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Apl
   ' *************************************************
   For i% = 1 To NbCon%
      For j% = 1 To NbVar%
         FenetreProLin.gridApl.Row = i%
         FenetreProLin.gridApl.Col = j%
         If FenetreProLin.gridApl.Text = "" Then
            FenetreProLin.gridApl.Text = 0
            Apl(i%, j%) = 0
         Else
            Apl(i%, j%) = CSng(FenetreProLin.gridApl.Text)
         End If
      Next j%
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Bpl
   ' *************************************************
   FenetreProLin.gridBpl.Col = 1
   For i% = 1 To NbCon%
      FenetreProLin.gridBpl.Row = i%
      If FenetreProLin.gridBpl.Text = "" Then
         FenetreProLin.gridBpl.Text = 0
         Bpl(i%) = 0
      Else
         Bpl(i%) = CSng(FenetreProLin.gridBpl.Text)
      End If
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Cpl
   ' *************************************************
   FenetreProLin.gridCpl.Col = 1
   For i% = 1 To NbVar%
      FenetreProLin.gridCpl.Row = i%
      If FenetreProLin.gridCpl.Text = "" Then
         FenetreProLin.gridCpl.Text = 0
         Cpl(i%) = 0
      Else
         Cpl(i%) = CSng(FenetreProLin.gridCpl.Text)
      End If
   Next i%
   ' *************************************************
   '    construction de A(i,j)
   ' *************************************************
   For i% = 1 To NbCon%
   Epl(i%) = 0
      For j% = 1 To NbCon%
         If j% = i% Then
            Apl(i%, NbVar% + j%) = 1
         Else
            Apl(i%, NbVar% + j%) = 0
         End If
      Next j%
      Apl(i%, NCpl%) = Bpl(i%)
   Next i%
   For j% = 1 To NbVar%
      Apl(NLpl%, j%) = Cpl(j%)
   Next j%
   For j% = NbVar% + 1 To NCpl%
      Apl(NLpl%, j%) = 0
   Next j%
   ' *************************************************
   '    algorithme du simplexe; méthode des tableaux
   ' *************************************************
   NbIter% = 0
   Do
      CP% = 1
      For jloc% = 2 To NCpl% - 1
         If Apl(NLpl%, jloc%) > Apl(NLpl%, CP%) Then CP% = jloc%
      Next jloc%
      If Apl(NLpl%, CP%) <= 0 Then Exit Do
      iloc% = 0
      Do
         iloc% = iloc% + 1
         If iloc% = NLpl% Then
            Beep
            MsgBox "Pas de convergence !", 48, "PROLIN"
            lblInfo.Caption = ""
            Exit Sub
         End If
         LP% = iloc%
         If Apl(iloc%, CP%) <> 0 Then
            Rloc0 = Apl(iloc%, NCpl%) / Apl(iloc%, CP%)
            If Rloc0 >= 0 Then
               Exit Do
            End If
         End If
      Loop
      Do
         iloc% = iloc% + 1
         If iloc% = NLpl% Then
            Exit Do
         End If
         If Apl(iloc%, CP%) <> 0 Then
            Rloc = Apl(iloc%, NCpl%) / Apl(iloc%, CP%)
            If Rloc >= 0 Then
               If Rloc < Rloc0 Then
                  Rloc0 = Rloc
                  LP% = iloc%
               End If
            End If
         End If
      Loop
      Pivot = Apl(LP%, CP%)
      NbIter% = NbIter% + 1
      For iloc% = 1 To NLpl%
         For jloc% = 1 To NCpl%
            If iloc% = LP% Then
               ANpl(iloc%, jloc%) = Apl(iloc%, jloc%) / Pivot
            ElseIf jloc% = CP% Then
               ANpl(iloc%, jloc%) = 0
            Else
               ANpl(iloc%, jloc%) = Apl(iloc%, jloc%) - Apl(iloc%, CP%) * Apl(LP%, jloc%) / Pivot
            End If
         Next jloc%
      Next iloc%
      For iloc% = 1 To NLpl%
         For jloc% = 1 To NCpl%
            Apl(iloc%, jloc%) = ANpl(iloc%, jloc%)
         Next jloc%
      Next iloc%
   Loop
   ' ------------- Fin des itérations -------------------------
   For iloc% = 1 To NbCon%
      jloc% = 0
      Do
         jloc% = jloc% + 1
         If jloc% = NCpl% Then
            Beep
            MsgBox "Pas de convergence !", 48, "PROLIN"
            lblInfo.Caption = ""
            Exit Sub
         End If
         If Apl(iloc%, jloc%) = 1 Then
            If jloc% > NbVar% Then
               Epl(jloc% - NbVar%) = Apl(iloc%, NCpl%)
            Else
               Xpl(jloc%) = Apl(iloc%, NCpl%)
            End If
            Exit Do
         End If
      Loop
   Next iloc%
   f = -Apl(NLpl%, NCpl%)
   ' *************************************************
   ' affichage du nombre d'itérations
   ' *************************************************
   lblNbIter.Caption = Format(NbIter%, "000")
   ' *************************************************
   ' affichage de la fonction objectif
   ' *************************************************
   lblValF.Caption = Format(f, "0.000")
   ' *************************************************
   ' affichage des éléments de Xpl
   ' *************************************************
   FenetreProLin.gridXpl.Col = 1
   For i% = 1 To NbVar%
      FenetreProLin.gridXpl.Row = i%
      FenetreProLin.gridXpl.Text = Format(Xpl(i%), "0.000")
   Next i%
   ' *************************************************
   ' affichage des éléments de Epl
   ' *************************************************
   FenetreProLin.gridEpl.Col = 1
   For i% = 1 To NbCon%
      FenetreProLin.gridEpl.Row = i%
      FenetreProLin.gridEpl.Text = Format(Epl(i%), "0.000")
   Next i%
   ' *************************************************
   ' affichage des coûts marginaux
   ' *************************************************
   FenetreProLin.gridCoutsMar.Col = 1
   For i% = 1 To NbCon%
      FenetreProLin.gridCoutsMar.Row = i%
      If Epl(i%) = 0 Then
         FenetreProLin.gridCoutsMar.Text = Format(-Apl(NLpl%, NbVar% + i%), "0.000")
      Else
         FenetreProLin.gridCoutsMar.Text = "0"
      End If
   Next i%
   ' ********************************************
   lblInfo.Caption = ""
   ' ********************************************
End Sub


Private Sub mnuCalculer_Click()
   Call Simplexe
End Sub


Private Sub mnuQuitter_Click()
   FenetreProLin.Hide
End Sub

Private Sub txtNbCon_Change()
   On Error Resume Next
   NbCon% = CInt(txtNbCon.Text)
   If Err.Number <> 0 Then
      If Val(txtNbCon.Text) = 0 Then
         NbCon% = 1
      Else
         MsgBox "le nombre de conditions est incorrect", 48, "POLYNOME"
         Exit Sub
      End If
   End If
   If NbCon% < 1 Then
      Beep
      MsgBox "le nombre de contraintes doit être supérieur ou égal à 1 !", 48, "PROLIN"
      NbCon% = 1
      txtNbCon.Text = "1"
   End If
   On Error GoTo 0
   '***************************************
   NLpl% = NbCon% + 1
   NCpl% = NbVar% + NbCon% + 1
   '***************************************
   gridApl.Rows = NbCon% + 1
   gridBpl.Rows = NbCon% + 1
   gridEpl.Rows = NbCon% + 1
   gridCoutsMar.Rows = NbCon% + 1
   ' ********************************
   ' renumérotation 1ère colonne Apl
   ' ********************************
   gridApl.Col = 0
   For i% = 1 To NbCon%
      gridApl.Row = i%
      gridApl.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère colonne Bpl
   ' ********************************
   gridBpl.Col = 0
   For i% = 1 To NbCon%
      gridBpl.Row = i%
      gridBpl.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère colonne Epl
   ' ********************************
   gridEpl.Col = 0
   For i% = 1 To NbCon%
      gridEpl.Row = i%
      gridEpl.Text = i%
   Next i%
   ' ************************************
   ' renumérotation 1ère colonne CoutsMar
   ' ************************************
   gridCoutsMar.Col = 0
   For i% = 1 To NbCon%
      gridCoutsMar.Row = i%
      gridCoutsMar.Text = i%
   Next i%
   ' ********************************
End Sub

Private Sub txtNbVar_Change()
   On Error Resume Next
   NbVar% = CInt(txtNbVar.Text)
   If Err.Number <> 0 Then
      If CInt(txtNbVar.Text) = 0 Then
         NbVar% = 1
      Else
         MsgBox "le nombre de variables est incorrect", 48, "POLYNOME"
         Exit Sub
      End If
   End If
   If NbVar% < 1 Then
      Beep
      MsgBox "le nombre de variables doit être supérieur ou égal à 1 !", 48, "PROLIN"
      NbVar% = 1
      txtNbVar.Text = "1"
   End If
   On Error GoTo 0
   '***************************************
   ' rappel : NLpl% = NbCon% + 1 ( n'a pas varié)
   NCpl% = NbVar% + NbCon% + 1
   '***************************************
   gridApl.Cols = NbVar% + 1
   gridCpl.Rows = NbVar% + 1
   gridXpl.Rows = NbVar% + 1
   '***************************************
   For i% = 0 To NbVar%
      FenetreProLin.gridApl.FixedAlignment(i%) = 2
      FenetreProLin.gridApl.ColWidth(i%) = 500
   Next i%
   ' ********************************
   ' renumérotation 1ère ligne Apl
   ' ********************************
   gridApl.Row = 0
   For i% = 1 To NbVar%
      gridApl.Col = i%
      gridApl.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère colonne Cpl
   ' ********************************
   gridCpl.Col = 0
   For i% = 1 To NbVar%
      gridCpl.Row = i%
      gridCpl.Text = i%
   Next i%
   ' ********************************
   ' renumérotation 1ère colonne Xpl
   ' ********************************
   gridXpl.Col = 0
   For i% = 1 To NbVar%
      gridXpl.Row = i%
      gridXpl.Text = i%
   Next i%
   ' ********************************
End Sub

