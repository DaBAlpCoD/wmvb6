VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetrePolynome 
   Caption         =   "Polynômes"
   ClientHeight    =   6480
   ClientLeft      =   300
   ClientTop       =   630
   ClientWidth     =   8775
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
   ScaleHeight     =   6480
   ScaleWidth      =   8775
   Begin MSFlexGridLib.MSFlexGrid GridRpol 
      Height          =   1935
      Left            =   6480
      TabIndex        =   40
      Top             =   4320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridQpol 
      Height          =   1935
      Left            =   4200
      TabIndex        =   39
      Top             =   4320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridP2pol 
      Height          =   1455
      Left            =   6480
      TabIndex        =   38
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridPPpol 
      Height          =   1935
      Left            =   1920
      TabIndex        =   37
      Top             =   4320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridRacines 
      Height          =   975
      Left            =   120
      TabIndex        =   36
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridPpol 
      Height          =   1455
      Left            =   4200
      TabIndex        =   35
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
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
   Begin VB.TextBox txtValX 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "1"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtValDegP2 
      Height          =   285
      Left            =   7560
      TabIndex        =   24
      Text            =   "2"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtValDegP 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "4"
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblValeur 
      Caption         =   "Valeur :"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblRacines 
      Caption         =   "Racines :"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblDefPol 
      Caption         =   "P(X) = Un X^n +...+ Up X^p +...+ U1 X + U0"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblCoefR 
      Caption         =   "Coefficients :"
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblValDegR 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7560
      TabIndex        =   28
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblDegR 
      Caption         =   "Degré :"
      Height          =   255
      Left            =   6480
      TabIndex        =   27
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblRpol 
      Caption         =   "Polynôme R reste de la division de P par P2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   26
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   2880
      Y2              =   6360
   End
   Begin VB.Label lblInfo 
      Height          =   495
      Left            =   2400
      TabIndex        =   25
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblValRacines 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label lblValP 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblPdeX 
      Alignment       =   1  'Right Justify
      Caption         =   "P(X) ="
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      Caption         =   "X ="
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   8640
      X2              =   8640
      Y1              =   120
      Y2              =   6360
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1800
      X2              =   8640
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4080
      X2              =   8640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   120
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1800
      X2              =   8640
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   6360
   End
   Begin VB.Label lblCoefQ 
      Caption         =   "Coefficients :"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblCoefPP 
      Caption         =   "Coefficients :"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblDegQ 
      Caption         =   "Degré :"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblValDegQ 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblValDegPP 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblDegPP 
      Caption         =   "Degré :"
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblDegP2 
      Caption         =   "Degré :"
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblCoefP2 
      Caption         =   "Coefficients :"
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblCoefP 
      Caption         =   "Coefficients :"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblQpol 
      Caption         =   "Polynôme Q quotient de P et P2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   17
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblPPpol 
      Caption         =   "Polynôme PP produit de P et P2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblP2pol 
      Caption         =   "Polynôme P2 :"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblDegP 
      Caption         =   "Degré : n ="
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblPpol 
      Caption         =   "Polynôme P :"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
         Begin VB.Menu mnuOuvrirP 
            Caption         =   "Polynôme P"
         End
         Begin VB.Menu mnuOuvrirP2 
            Caption         =   "Polynôme P2"
         End
      End
      Begin VB.Menu mnuEnreg 
         Caption         =   "&Enregistrer..."
         Begin VB.Menu mnuEnregP 
            Caption         =   "Polynôme P"
         End
         Begin VB.Menu mnuEnregP2 
            Caption         =   "Polynôme P2"
         End
         Begin VB.Menu mnuEnregPP 
            Caption         =   "Polynôme PP"
         End
         Begin VB.Menu mnuEnregQ 
            Caption         =   "Polynôme Q"
         End
         Begin VB.Menu mnuEnregR 
            Caption         =   "Poynôme R"
         End
      End
   End
   Begin VB.Menu mnuCalculer 
      Caption         =   "&Calculer..."
      Begin VB.Menu mnuRacines 
         Caption         =   "&Racines"
         Begin VB.Menu mnuRacinesExactes 
            Caption         =   "Méthodes exactes (n<5)"
         End
         Begin VB.Menu mnuRacinesMatComp 
            Caption         =   "Méthode de la matrice compagne (n>2)"
         End
         Begin VB.Menu mnuRacinesBair 
            Caption         =   "Méthode de BAIRSTOW"
         End
      End
      Begin VB.Menu mnuValeur 
         Caption         =   "&Valeur"
      End
      Begin VB.Menu mnuProduit 
         Caption         =   "&Produit PxP2"
      End
      Begin VB.Menu mnuQuotient 
         Caption         =   "&Quotient P/P2"
      End
   End
   Begin VB.Menu mnuTracer 
      Caption         =   "&Tracer..."
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetrePolynome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xpol As Single
Dim NomPol$








Private Sub Form_Activate()
   ' ********************************
   ' actualisation des 6 grilles
   ' ********************************
   FenetrePolynome.gridPpol.Rows = DegPpol% + 2
   FenetrePolynome.gridPpol.Cols = 2
   FenetrePolynome.txtValDegP2.Text = Format(DegP2pol%, "0")
   FenetrePolynome.gridP2pol.Rows = DegP2pol% + 2
   FenetrePolynome.gridP2pol.Cols = 2
   FenetrePolynome.lblValDegPP.Caption = ""
   FenetrePolynome.gridPPpol.Rows = DegPPpol% + 2
   FenetrePolynome.gridPPpol.Cols = 2
   FenetrePolynome.lblValDegQ.Caption = ""
   FenetrePolynome.gridQpol.Rows = DegQpol% + 2
   FenetrePolynome.gridQpol.Cols = 2
   FenetrePolynome.lblValDegR.Caption = ""
   FenetrePolynome.gridRpol.Rows = DegRpol% + 2
   FenetrePolynome.gridRpol.Cols = 2
   FenetrePolynome.gridRacines.Rows = DegPpol% + 1
   FenetrePolynome.gridRacines.Cols = 3
   ' ********************************
   ' numérotation 1ères colonnes
   ' ********************************
   FenetrePolynome.gridPpol.Col = 0
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      FenetrePolynome.gridPpol.Text = Format(i%, "0")
   Next i%
   '************************************
     FenetrePolynome.gridP2pol.Col = 0
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      FenetrePolynome.gridP2pol.Text = Format(i%, "0")
   Next i%
   '************************************
   FenetrePolynome.gridPPpol.Col = 0
   For i% = 0 To DegPPpol%
      FenetrePolynome.gridPPpol.Row = i% + 1
      FenetrePolynome.gridPPpol.Text = Format(i%, "0")
   Next i%
  '************************************
     FenetrePolynome.gridQpol.Col = 0
   For i% = 0 To DegQpol%
      FenetrePolynome.gridQpol.Row = i% + 1
      FenetrePolynome.gridQpol.Text = Format(i%, "0")
   Next i%
   '************************************
     FenetrePolynome.gridRpol.Col = 0
   For i% = 0 To DegRpol%
      FenetrePolynome.gridRpol.Row = i% + 1
      FenetrePolynome.gridRpol.Text = Format(i%, "0")
   Next i%
   '************************************
   FenetrePolynome.gridRacines.Col = 0
   For i% = 1 To DegPpol%
      FenetrePolynome.gridRacines.Row = i%
      FenetrePolynome.gridRacines.Text = Format(i%, "0")
   Next i%
   '***************************************
   txtValX.Text = Format(Xpol, "0.000")
   ' ********************************
   ' placement des valeurs
   ' ********************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      FenetrePolynome.gridPpol.Text = Format(Ppol(i%), "0.000")
   Next i%
   FenetrePolynome.gridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      FenetrePolynome.gridP2pol.Text = Format(P2pol(i%), "0.000")
   Next i%
   ' ****************************************
   ' ne pas placer cette ligne avant le placement des
   ' valeurs de Ppol car elle provoque un évènement
   ' txtValDegP_Change qui provoque un ReDim Ppol et
   ' remet les valeurs de Ppol à 0
   FenetrePolynome.txtValDegP.Text = Format(DegPpol%, "0")
   ' ******************************************
End Sub

Private Sub Form_Deactivate()
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de P2pol
   ' *************************************************
   FenetrePolynome.gridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      If FenetrePolynome.gridP2pol.Text = "" Then
         FenetrePolynome.gridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.gridP2pol.Text)
      End If
   Next i%
   ' *************************************************
End Sub


Private Sub Form_Load()
   '***************************************
   DegPpol% = 4
   DegP2pol% = 2
   DegPPpol% = DegPpol% + DegP2pol%
   DegQpol% = DegPpol% - DegP2pol%
   DegRpol% = DegP2pol% - 1
   '***************************************
   ReDim Ppol(0 To DegPpol%)
   ReDim P2pol(0 To DegP2pol%)
   ReDim PPpol(0 To DegPPpol%)
   ReDim Qpol(0 To DegQpol%)
   ReDim Rpol(0 To DegRpol%)
   ' ********************************
   ' mise en place des grilles
   ' ********************************
   FenetrePolynome.txtValDegP.Text = Format(DegPpol%, "0")
   FenetrePolynome.gridPpol.Rows = DegPpol% + 2
   FenetrePolynome.gridPpol.Cols = 2
   FenetrePolynome.txtValDegP2.Text = Format(DegP2pol%, "0")
   FenetrePolynome.gridP2pol.Rows = DegP2pol% + 2
   FenetrePolynome.gridP2pol.Cols = 2
   FenetrePolynome.lblValDegPP.Caption = ""
   FenetrePolynome.gridPPpol.Rows = DegPPpol% + 2
   FenetrePolynome.gridPPpol.Cols = 2
   FenetrePolynome.lblValDegQ.Caption = ""
   FenetrePolynome.gridQpol.Rows = DegQpol% + 2
   FenetrePolynome.gridQpol.Cols = 2
   FenetrePolynome.lblValDegR.Caption = ""
   FenetrePolynome.gridRpol.Rows = DegRpol% + 2
   FenetrePolynome.gridRpol.Cols = 2
   FenetrePolynome.gridRacines.Rows = DegPpol% + 1
   FenetrePolynome.gridRacines.Cols = 3
   ' ********************************
   ' largeur des colonnes et
   ' mentions dans les 1ères lignes
   ' ********************************
   FenetrePolynome.gridPpol.FixedAlignment(0) = 2
   FenetrePolynome.gridPpol.FixedAlignment(1) = 2
   FenetrePolynome.gridPpol.ColWidth(0) = 500
   FenetrePolynome.gridPpol.ColWidth(1) = 1200
   FenetrePolynome.gridPpol.Row = 0
   FenetrePolynome.gridPpol.Col = 0
   FenetrePolynome.gridPpol.Text = "p"
   FenetrePolynome.gridPpol.Col = 1
   FenetrePolynome.gridPpol.Text = "Up"
   '***************************************
   FenetrePolynome.gridP2pol.FixedAlignment(0) = 2
   FenetrePolynome.gridP2pol.FixedAlignment(1) = 2
   FenetrePolynome.gridP2pol.ColWidth(0) = 500
   FenetrePolynome.gridP2pol.ColWidth(1) = 1200
   FenetrePolynome.gridP2pol.Row = 0
   FenetrePolynome.gridP2pol.Col = 0
   FenetrePolynome.gridP2pol.Text = "p"
   FenetrePolynome.gridP2pol.Col = 1
   FenetrePolynome.gridP2pol.Text = "Up"
   '***************************************
   FenetrePolynome.gridPPpol.FixedAlignment(0) = 2
   FenetrePolynome.gridPPpol.FixedAlignment(1) = 2
   FenetrePolynome.gridPPpol.ColWidth(0) = 500
   FenetrePolynome.gridPPpol.ColWidth(1) = 1200
   FenetrePolynome.gridPPpol.Row = 0
   FenetrePolynome.gridPPpol.Col = 0
   FenetrePolynome.gridPPpol.Text = "p"
   FenetrePolynome.gridPPpol.Col = 1
   FenetrePolynome.gridPPpol.Text = "Up"
   '***************************************
   FenetrePolynome.gridQpol.FixedAlignment(0) = 2
   FenetrePolynome.gridQpol.FixedAlignment(1) = 2
   FenetrePolynome.gridQpol.ColWidth(0) = 500
   FenetrePolynome.gridQpol.ColWidth(1) = 1200
   FenetrePolynome.gridQpol.Row = 0
   FenetrePolynome.gridQpol.Col = 0
   FenetrePolynome.gridQpol.Text = "p"
   FenetrePolynome.gridQpol.Col = 1
   FenetrePolynome.gridQpol.Text = "Up"
   '***************************************
   FenetrePolynome.gridRpol.FixedAlignment(0) = 2
   FenetrePolynome.gridRpol.FixedAlignment(1) = 2
   FenetrePolynome.gridRpol.ColWidth(0) = 500
   FenetrePolynome.gridRpol.ColWidth(1) = 1200
   FenetrePolynome.gridRpol.Row = 0
   FenetrePolynome.gridRpol.Col = 0
   FenetrePolynome.gridRpol.Text = "p"
   FenetrePolynome.gridRpol.Col = 1
   FenetrePolynome.gridRpol.Text = "Up"
   '***************************************
   FenetrePolynome.gridRacines.FixedAlignment(0) = 2
   FenetrePolynome.gridRacines.FixedAlignment(1) = 2
   FenetrePolynome.gridRacines.FixedAlignment(2) = 2
   FenetrePolynome.gridRacines.ColWidth(0) = 500
   FenetrePolynome.gridRacines.ColWidth(1) = 1400
   FenetrePolynome.gridRacines.ColWidth(2) = 1400
   FenetrePolynome.gridRacines.Row = 0
   FenetrePolynome.gridRacines.Col = 0
   FenetrePolynome.gridRacines.Text = "N°"
   FenetrePolynome.gridRacines.Col = 1
   FenetrePolynome.gridRacines.Text = "Partie réelle"
   FenetrePolynome.gridRacines.Col = 2
   FenetrePolynome.gridRacines.Text = "et imaginaire"
   ' ********************************
   ' numérotation 1ères colonnes
   ' ********************************
   FenetrePolynome.gridPpol.Col = 0
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      FenetrePolynome.gridPpol.Text = Format(i%, "0")
   Next i%
   '************************************
   FenetrePolynome.gridP2pol.Col = 0
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      FenetrePolynome.gridP2pol.Text = Format(i%, "0")
   Next i%
   '************************************
   FenetrePolynome.gridPPpol.Col = 0
   FenetrePolynome.gridPPpol.Row = 1
   FenetrePolynome.gridPPpol.Text = "0"
   '************************************
   FenetrePolynome.gridQpol.Col = 0
   FenetrePolynome.gridQpol.Row = 1
   FenetrePolynome.gridQpol.Text = "0"
   '************************************
   FenetrePolynome.gridRpol.Col = 0
   FenetrePolynome.gridRpol.Row = 1
   FenetrePolynome.gridRpol.Text = "0"
   ' ********************************
   FenetrePolynome.gridRacines.Col = 0
   For i% = 1 To DegPpol%
      FenetrePolynome.gridRacines.Row = i%
      FenetrePolynome.gridRacines.Text = Format(i%, "0")
   Next i%
   ' ********************************
   '       Valeurs par défaut
   ' ********************************
   Ppol(0) = -2
   Ppol(1) = 2
   Ppol(2) = 1
   Ppol(3) = -2
   Ppol(4) = 1
   '***************************************
   P2pol(0) = -6
   P2pol(1) = -1
   P2pol(2) = 1
   '***************************************
   Xpol = 1
   FenetrePolynome.txtValX.Text = Format(Xpol, "0.000")
   ' ********************************
   ' placement des valeurs par défaut
   ' ********************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      FenetrePolynome.gridPpol.Text = Format(Ppol(i%), "0.000")
   Next i%
   FenetrePolynome.gridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      FenetrePolynome.gridP2pol.Text = Format(P2pol(i%), "0.000")
   Next i%
   ' ****************************************
End Sub

Private Sub gridP2pol_KeyPress(KeyAscii As Integer)
   '-------------------------------------------
   ' Remises à zéro
   '-------------------------------------------
   ' ********************************
   ' placement des 'blancs'
   ' ********************************
   FenetrePolynome.gridPPpol.Col = 1
   For i% = 0 To DegPPpol%
      FenetrePolynome.gridPPpol.Row = i% + 1
      FenetrePolynome.gridPPpol.Text = ""
   Next i%
   ' ********************************
   FenetrePolynome.gridQpol.Col = 1
   For i% = 0 To DegQpol%
      FenetrePolynome.gridQpol.Row = i% + 1
      FenetrePolynome.gridQpol.Text = ""
   Next i%
   ' ****************************************
   FenetrePolynome.gridRpol.Col = 1
   For i% = 0 To DegRpol%
      FenetrePolynome.gridRpol.Row = i% + 1
      FenetrePolynome.gridRpol.Text = ""
   Next i%
   ' ****************************************
   '-------------------------------------------
   ' Ecriture
   '-------------------------------------------
   ElePText$ = gridP2pol.Text
   Select Case KeyAscii
   Case 32 To 168
      ElePCar$ = Chr(KeyAscii)
      ElePText$ = ElePText$ & ElePCar$
      gridP2pol.Text = ElePText$
   Case 8
      If Len(gridP2pol.Text) > 0 Then
         ElePText$ = Left$(ElePText$, Len(ElePText$) - 1)
         gridP2pol.Text = ElePText$
      Else
         Beep
      End If
   End Select
End Sub

Private Sub gridPpol_KeyPress(KeyAscii As Integer)
   '-------------------------------------------
   ' Remises à zéro
   '-------------------------------------------
   lblValRacines.Caption = ""
   lblValP.Caption = ""
   ' ********************************
   ' placement des 'blancs'
   ' ********************************
   FenetrePolynome.gridPPpol.Col = 1
   For i% = 0 To DegPPpol%
      FenetrePolynome.gridPPpol.Row = i% + 1
      FenetrePolynome.gridPPpol.Text = ""
   Next i%
   ' ********************************
   FenetrePolynome.gridQpol.Col = 1
   For i% = 0 To DegQpol%
      FenetrePolynome.gridQpol.Row = i% + 1
      FenetrePolynome.gridQpol.Text = ""
   Next i%
   ' ****************************************
   FenetrePolynome.gridRpol.Col = 1
   For i% = 0 To DegRpol%
      FenetrePolynome.gridRpol.Row = i% + 1
      FenetrePolynome.gridRpol.Text = ""
   Next i%
   ' ****************************************
   For jloc% = 1 To 2
      FenetrePolynome.gridRacines.Col = jloc%
      For iloc% = 1 To DegPpol%
         FenetrePolynome.gridRacines.Row = iloc%
         FenetrePolynome.gridRacines.Text = ""
      Next iloc%
   Next jloc%
   ' ********************************
   '-------------------------------------------
   ' Ecriture
   '-------------------------------------------
   ElePText$ = gridPpol.Text
   Select Case KeyAscii
   Case 32 To 168
      ElePCar$ = Chr(KeyAscii)
      ElePText$ = ElePText$ & ElePCar$
      gridPpol.Text = ElePText$
   Case 8
      If Len(gridPpol.Text) > 0 Then
         ElePText$ = Left$(ElePText$, Len(ElePText$) - 1)
         gridPpol.Text = ElePText$
      Else
         Beep
      End If
   End Select
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
   FenetrePolynome.gridPPpol.Rows = DegPPpol% + 2
   FenetrePolynome.gridPPpol.Cols = 2
   ' ********************************
   ' numérotation 1ère colonnes
   ' ********************************
   FenetrePolynome.gridPPpol.Col = 0
   For i% = 0 To DegPPpol%
      FenetrePolynome.gridPPpol.Row = i% + 1
      FenetrePolynome.gridPPpol.Text = Format(i%, "0")
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de P2pol
   ' *************************************************
   FenetrePolynome.gridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      If FenetrePolynome.gridP2pol.Text = "" Then
         FenetrePolynome.gridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.gridP2pol.Text)
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
   FenetrePolynome.gridPPpol.Col = 1
   For i% = 0 To DegPPpol%
      FenetrePolynome.gridPPpol.Row = i% + 1
      FenetrePolynome.gridPPpol.Text = Format(PPpol(i%), "0.000")
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
   FenetrePolynome.gridQpol.Rows = DegQpol% + 2
   FenetrePolynome.gridQpol.Cols = 2
   FenetrePolynome.lblValDegR.Caption = Format(DegRpol%, "0")
   FenetrePolynome.gridRpol.Rows = DegRpol% + 2
   FenetrePolynome.gridRpol.Cols = 2
   ' ********************************
   ' numérotation 1ères colonnes
   ' ********************************
   FenetrePolynome.gridQpol.Col = 0
   For i% = 0 To DegQpol%
      FenetrePolynome.gridQpol.Row = i% + 1
      FenetrePolynome.gridQpol.Text = Format(i%, "0")
   Next i%
   FenetrePolynome.gridRpol.Col = 0
   For i% = 0 To DegRpol%
      FenetrePolynome.gridRpol.Row = i% + 1
      FenetrePolynome.gridRpol.Text = Format(i%, "0")
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de P2pol
   ' *************************************************
   FenetrePolynome.gridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      If FenetrePolynome.gridP2pol.Text = "" Then
         FenetrePolynome.gridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.gridP2pol.Text)
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
   FenetrePolynome.gridQpol.Col = 1
   For i% = 0 To DegQpol%
      FenetrePolynome.gridQpol.Row = i% + 1
      FenetrePolynome.gridQpol.Text = Format(Qpol(i%), "0.000")
   Next i%
   FenetrePolynome.gridRpol.Col = 1
   For i% = 0 To DegRpol%
      FenetrePolynome.gridRpol.Row = i% + 1
      FenetrePolynome.gridRpol.Text = Format(Rpol(i%), "0.000")
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


Private Sub mnuProduit_Click()
   Call ProduitPol
End Sub

Private Sub mnuQuitter_Click()
   FenetrePolynome.Hide
End Sub


Private Sub mnuQuotient_Click()
   Call QuotientPol
End Sub


Private Sub mnuRacinesBair_Click()
   Call RacinesBairstow
End Sub

Private Sub mnuRacinesExactes_Click()
   Call RacinesExactes
End Sub



Private Sub mnuRacinesMatComp_Click()
   Call RacinesMatComp
End Sub

Private Sub mnuTracer_Click()
   FenetrePolynome.Hide
   FenetreDefPolynome.Show
End Sub



Private Sub mnuValeur_Click()
   Xpol = CSng(txtValX.Text)
   If Err.Number <> 0 Then
      MsgBox "Xpol est incorrect", 48, "Polynome"
      Exit Sub
   End If
   Call Valeur
End Sub


Private Sub txtValDegP_Change()
   On Error Resume Next
   DegPpol% = CInt(txtValDegP.Text)
   If Err.Number <> 0 Then
      If Val(txtValDegP.Text) = 0 Then
         DegPpol% = 0
      Else
         MsgBox "le degré de P est incorrect", 48, "POLYNOME"
         Exit Sub
      End If
   End If
   On Error GoTo 0
   If DegPpol% < 0 Then
      Beep
      MsgBox "le degré de P doit être supérieur ou égal à 0 !", 48, "POLYNOME"
      DegPpol% = 0
      txtValDegP.Text = "0"
   End If
   lblValRacines.Caption = ""
   lblValP.Caption = ""
   lblValDegPP.Caption = ""
   lblValDegQ.Caption = ""
   lblValDegR.Caption = ""
   DegPPpol% = 0
   DegQpol% = 0
   DegRpol% = 0
   gridPpol.Rows = DegPpol% + 2
   gridPPpol.Rows = 2
   gridQpol.Rows = 2
   gridRpol.Rows = 2
   gridPPpol.Col = 1
   gridPPpol.Row = 1
   gridPPpol.Text = ""
   gridQpol.Col = 1
   gridQpol.Row = 1
   gridQpol.Text = ""
   gridRpol.Col = 1
   gridRpol.Row = 1
   gridRpol.Text = ""
   gridRacines.Rows = DegPpol% + 1
   '***************************************
   ReDim Ppol(0 To DegPpol%)
   '***************************************
   ' renumérotation 1ère colonne Ppol
   ' ********************************
   FenetrePolynome.gridPpol.Col = 0
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      FenetrePolynome.gridPpol.Text = Format(i%, "0")
   Next i%
   '***************************************
   ' renumérotation 1ère colonne gridRacines
   ' et effacement valeurs racines
   ' ********************************
   For i% = 1 To DegPpol%
      FenetrePolynome.gridRacines.Row = i%
      FenetrePolynome.gridRacines.Col = 0
      FenetrePolynome.gridRacines.Text = Format(i%, "0")
      FenetrePolynome.gridRacines.Col = 1
      FenetrePolynome.gridRacines.Text = ""
      FenetrePolynome.gridRacines.Col = 2
      FenetrePolynome.gridRacines.Text = ""
   Next i%
End Sub

Private Sub txtValDegP2_Change()
   On Error Resume Next
   DegP2pol% = CInt(txtValDegP2.Text)
   If Err.Number <> 0 Then
      If CInt(txtValDegP2.Text) = 0 Then
         DegP2pol% = 0
      Else
         MsgBox "le degré de P2 est incorrect", 48, "POLYNOME"
         Exit Sub
      End If
   End If
   On Error GoTo 0
   If DegP2pol% < 0 Then
      Beep
      MsgBox "le degré de P2 doit être supérieur ou égal à 0 !", 48, "POLYNOME"
      DegP2pol% = 0
      txtValDegP2.Text = "0"
   End If
   lblValDegPP.Caption = ""
   lblValDegQ.Caption = ""
   lblValDegR.Caption = ""
   DegPPpol% = 0
   DegQpol% = 0
   DegRpol% = 0
   gridP2pol.Rows = DegP2pol% + 2
   gridPPpol.Rows = 2
   gridQpol.Rows = 2
   gridRpol.Rows = 2
   FenetrePolynome.gridPPpol.Col = 1
   FenetrePolynome.gridPPpol.Row = 1
   FenetrePolynome.gridPPpol.Text = ""
   FenetrePolynome.gridQpol.Col = 1
   FenetrePolynome.gridQpol.Row = 1
   FenetrePolynome.gridQpol.Text = ""
   FenetrePolynome.gridRpol.Col = 1
   FenetrePolynome.gridRpol.Row = 1
   FenetrePolynome.gridRpol.Text = ""
   '***************************************
   ReDim P2pol(0 To DegP2pol%)
   '***************************************
   ' renumérotation 1ères colonnes
   ' ********************************
   FenetrePolynome.gridP2pol.Col = 0
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      FenetrePolynome.gridP2pol.Text = Format(i%, "0")
   Next i%
   ' ******************************************
End Sub

Private Sub txtValX_Change()
   On Error Resume Next
   Xpol = CSng(txtValX.Text)
   lblValP.Caption = ""
   On Error GoTo 0
End Sub


Public Sub Racines()
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
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
      FenetrePolynome.gridRacines.Row = 1
      FenetrePolynome.gridRacines.Col = 1
      FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
      FenetrePolynome.gridRacines.Col = 2
      FenetrePolynome.gridRacines.Text = "0"
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
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         ImPol = RacDis / Ppol(2) / 2
         XPol1 = ReelPol - ImPol
         XPol2 = ReelPol + ImPol
         TexteRacines$ = " Deux racines réelles :"
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      Else
         RacDis = Sqr(-Dis)
         ImPol = Abs(RacDis / Ppol(2) / 2)
         TexteRacines$ = " Deux racines complexes conjuguées :"
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
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
            FenetrePolynome.gridRacines.Row = 1
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 2
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 3
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
         Else
            XPol1 = 2 * Qloc / Ploc - Ppol(2) / Ppol(3) / 3
            XPol2 = -Qloc / Ploc - Ppol(2) / Ppol(3) / 3
            TexteRacines$ = TexteRacines$ & " Deux racines réelles : une simple et une double :"
            FenetrePolynome.gridRacines.Row = 1
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 2
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 3
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
         End If
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         GAloc = Sgn(-Qloc + RacDis) * Abs(-Qloc + RacDis) ^ (1 / 3)
         GBloc = Sgn(-Qloc - RacDis) * Abs(-Qloc - RacDis) ^ (1 / 3)
         ReelPol1 = GAloc + GBloc - Ppol(2) / 3 / Ppol(3)
         ReelPol2 = (-GAloc - GBloc) / 2 - Ppol(2) / 3 / Ppol(3)
         ImPol = Abs(Sqr(3) / 2 * (GAloc - GBloc))
         TexteRacines$ = TexteRacines$ & " Trois racines : une réelle et deux complexes conjuguées"
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
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
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol3, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
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
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      ElseIf DMloc >= 0 Then
         ImPol = Sqr(DMloc) / 2
         XPol2 = ReelPol - ImPol
         XPol3 = ReelPol + ImPol
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol3, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      Else
         ImPol = Sqr(-DMloc) / 2
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
      End If
      BBloc = Ppol(3) / Ppol(4) / 2 - APloc
      CBloc = Zloc - BPloc
      DMloc = BBloc * BBloc - 4 * CBloc
      ReelPol = -BBloc / 2
      If Abs(DMloc) < 0.000001 Then
         XPol1 = ReelPol
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 4
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      ElseIf DMloc >= 0 Then
         ImPol = Sqr(DMloc) / 2
         XPol2 = ReelPol - ImPol
         XPol3 = ReelPol + ImPol
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 4
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol3, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      Else
         ImPol = Sqr(-DMloc) / 2
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 4
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
      End If
      lblValRacines.Caption = TexteRacines$
   Case Is > 4
      Call ZerosPolBairstow
      If Erreur = True Then
         TexteRacine$ = "Impossible de calculer les racines !"
      Else
         TexteRacines$ = "      Calcul par la méthode de BAIRSTOW" & Chr$(13) & Chr$(13)
         For iloc% = 1 To DegPpol%
            FenetrePolynome.gridRacines.Row = iloc%
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(RacineR(iloc%), "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = Format(RacineI(iloc%), "0.000")
         Next iloc%
      End If
      lblValRacines.Caption = TexteRacines$
   End Select
End Sub

Public Sub Valeur()
   ' ********************************************
   '               ReDims
   ' ********************************************
   ReDim Ppol(0 To DegPpol%)
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
   px = 0
   Xpol = CSng(txtValX.Text)
   For i% = 0 To DegPpol%
      px = px + Ppol(i%) * Xpol ^ i%
   Next i%
   lblValP.Caption = Format(px, "0.000")
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
      FenetrePolynome.gridPpol.Col = 1
      For i% = 0 To DegPolLoc%
         FenetrePolynome.gridPpol.Row = i% + 1
         FenetrePolynome.gridPpol.Text = Format(Ppol(i%), "0.000")
      Next i%
   ElseIf NomPol$ = "P2" Then
      FenetrePolynome.gridP2pol.Col = 1
      For i% = 0 To DegPolLoc%
         FenetrePolynome.gridP2pol.Row = i% + 1
         FenetrePolynome.gridP2pol.Text = Format(P2pol(i%), "0.000")
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
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.gridP2pol.Col = 1
   For i% = 0 To DegP2pol%
      FenetrePolynome.gridP2pol.Row = i% + 1
      If FenetrePolynome.gridP2pol.Text = "" Then
         FenetrePolynome.gridP2pol.Text = "0"
         P2pol(i%) = 0
      Else
         P2pol(i%) = CSng(FenetrePolynome.gridP2pol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.gridPPpol.Col = 1
   For i% = 0 To DegPPpol%
      FenetrePolynome.gridPPpol.Row = i% + 1
      If FenetrePolynome.gridPPpol.Text = "" Then
         FenetrePolynome.gridPPpol.Text = 0
         PPpol(i%) = 0
      Else
         PPpol(i%) = CSng(FenetrePolynome.gridPPpol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.gridQpol.Col = 1
   For i% = 0 To DegQpol%
      FenetrePolynome.gridQpol.Row = i% + 1
      If FenetrePolynome.gridQpol.Text = "" Then
         FenetrePolynome.gridQpol.Text = "0"
         Qpol(i%) = 0
      Else
         Qpol(i%) = CSng(FenetrePolynome.gridQpol.Text)
      End If
   Next i%
   ' *************************************************
   FenetrePolynome.gridRpol.Col = 1
   For i% = 0 To DegRpol%
      FenetrePolynome.gridRpol.Row = i% + 1
      If FenetrePolynome.gridRpol.Text = "" Then
         FenetrePolynome.gridRpol.Text = "0"
         Rpol(i%) = 0
      Else
         Rpol(i%) = CSng(FenetrePolynome.gridRpol.Text)
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
      gridPpol.Col = 1
      For i% = 0 To DegPpol%
         gridPpol.Row = i% + 1
         Write #1, gridPpol.Text
      Next i%
   ElseIf NomPol$ = "P2" Then
      Write #1, DegP2pol%
      gridP2pol.Col = 1
      For i% = 0 To DegP2pol%
         gridP2pol.Row = i% + 1
         Write #1, gridP2pol.Text
      Next i%
   ElseIf NomPol$ = "PP" Then
      Write #1, DegPPpol%
      gridPPpol.Col = 1
      For i% = 0 To DegPPpol%
         gridPPpol.Row = i% + 1
         Write #1, gridPPpol.Text
      Next i%
   ElseIf NomPol$ = "Q" Then
      Write #1, DegQpol%
      gridQpol.Col = 1
      For i% = 0 To DegQpol%
         gridQpol.Row = i% + 1
         Write #1, gridQpol.Text
      Next i%
   ElseIf NomPol$ = "R" Then
      Write #1, DegRpol%
      gridRpol.Col = 1
      For i% = 0 To DegRpol%
         gridRpol.Row = i% + 1
         Write #1, gridRpol.Text
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
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
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
      FenetrePolynome.gridRacines.Row = 1
      FenetrePolynome.gridRacines.Col = 1
      FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
      FenetrePolynome.gridRacines.Col = 2
      FenetrePolynome.gridRacines.Text = "0"
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
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         ImPol = RacDis / Ppol(2) / 2
         XPol1 = ReelPol - ImPol
         XPol2 = ReelPol + ImPol
         TexteRacines$ = " Deux racines réelles :"
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      Else
         RacDis = Sqr(-Dis)
         ImPol = Abs(RacDis / Ppol(2) / 2)
         TexteRacines$ = " Deux racines complexes conjuguées :"
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
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
            FenetrePolynome.gridRacines.Row = 1
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 2
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 3
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
         Else
            XPol1 = 2 * Qloc / Ploc - Ppol(2) / Ppol(3) / 3
            XPol2 = -Qloc / Ploc - Ppol(2) / Ppol(3) / 3
            TexteRacines$ = TexteRacines$ & " Deux racines réelles : une simple et une double :"
            FenetrePolynome.gridRacines.Row = 1
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 2
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
            FenetrePolynome.gridRacines.Row = 3
            FenetrePolynome.gridRacines.Col = 1
            FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
            FenetrePolynome.gridRacines.Col = 2
            FenetrePolynome.gridRacines.Text = "0"
         End If
      ElseIf Dis > 0 Then
         RacDis = Sqr(Dis)
         GAloc = Sgn(-Qloc + RacDis) * Abs(-Qloc + RacDis) ^ (1 / 3)
         GBloc = Sgn(-Qloc - RacDis) * Abs(-Qloc - RacDis) ^ (1 / 3)
         ReelPol1 = GAloc + GBloc - Ppol(2) / 3 / Ppol(3)
         ReelPol2 = (-GAloc - GBloc) / 2 - Ppol(2) / 3 / Ppol(3)
         ImPol = Abs(Sqr(3) / 2 * (GAloc - GBloc))
         TexteRacines$ = TexteRacines$ & " Trois racines : une réelle et deux complexes conjuguées"
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
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
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol3, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
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
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      ElseIf DMloc >= 0 Then
         ImPol = Sqr(DMloc) / 2
         XPol2 = ReelPol - ImPol
         XPol3 = ReelPol + ImPol
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol3, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      Else
         ImPol = Sqr(-DMloc) / 2
         FenetrePolynome.gridRacines.Row = 1
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 2
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
      End If
      BBloc = Ppol(3) / Ppol(4) / 2 - APloc
      CBloc = Zloc - BPloc
      DMloc = BBloc * BBloc - 4 * CBloc
      ReelPol = -BBloc / 2
      If Abs(DMloc) < 0.000001 Then
         XPol1 = ReelPol
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 4
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol1, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      ElseIf DMloc >= 0 Then
         ImPol = Sqr(DMloc) / 2
         XPol2 = ReelPol - ImPol
         XPol3 = ReelPol + ImPol
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol2, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
         FenetrePolynome.gridRacines.Row = 4
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(XPol3, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = "0"
      Else
         ImPol = Sqr(-DMloc) / 2
         FenetrePolynome.gridRacines.Row = 3
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(-ImPol, "0.000")
         FenetrePolynome.gridRacines.Row = 4
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(ReelPol, "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(ImPol, "0.000")
      End If
      lblValRacines.Caption = TexteRacines$
   Case Is > 4
      Message$ = "Impossible de calculer de manière exacte"
      Message$ = Message$ & Chr$(13) & "les racines d'un polynôme de degré supérieur à 4 !"
      MsgBox Message$, 48, "POLYNOME"
   End Select
End Sub

Public Sub RacinesBairstow()
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
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
   End Select
   Call ZerosPolBairstow
   If Erreur = True Then
      TexteRacine$ = "Impossible de calculer les racines !"
   Else
      TexteRacines$ = "      Calcul par la méthode de BAIRSTOW" & Chr$(13) & Chr$(13)
      For iloc% = 1 To DegPpol%
         FenetrePolynome.gridRacines.Row = iloc%
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(RacineR(iloc%), "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(RacineI(iloc%), "0.000")
      Next iloc%
   End If
   lblValRacines.Caption = TexteRacines$
End Sub

Public Sub RacinesMatComp()
   ' *************************************************
   ' affectation de leurs valeurs aux éléments de Ppol
   ' *************************************************
   FenetrePolynome.gridPpol.Col = 1
   For i% = 0 To DegPpol%
      FenetrePolynome.gridPpol.Row = i% + 1
      If FenetrePolynome.gridPpol.Text = "" Then
         FenetrePolynome.gridPpol.Text = "0"
         Ppol(i%) = 0
      Else
         Ppol(i%) = CSng(FenetrePolynome.gridPpol.Text)
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
   End Select
   Call ZerosPolMatComp
   If Erreur = True Then
      TexteRacine$ = "Impossible de calculer les racines !"
   Else
      TexteRacines$ = "Calcul par la méthode de la matrice compagne" & Chr$(13) & Chr$(13)
      For iloc% = 1 To DegPpol%
         FenetrePolynome.gridRacines.Row = iloc%
         FenetrePolynome.gridRacines.Col = 1
         FenetrePolynome.gridRacines.Text = Format(RacineR(iloc%), "0.000")
         FenetrePolynome.gridRacines.Col = 2
         FenetrePolynome.gridRacines.Text = Format(RacineI(iloc%), "0.000")
      Next iloc%
   End If
   lblValRacines.Caption = TexteRacines$
End Sub
