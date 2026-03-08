VERSION 5.00
Begin VB.Form FenetrePPCMPGCD 
   Caption         =   "PPCM et PGCD de deux nombres entiers"
   ClientHeight    =   4935
   ClientLeft      =   735
   ClientTop       =   795
   ClientWidth     =   8280
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
   ScaleHeight     =   4935
   ScaleWidth      =   8280
   Begin VB.TextBox TextPGCD 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   3360
      Width           =   5535
   End
   Begin VB.TextBox TextPPCM 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox TextN2 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox TextN1 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   5535
   End
   Begin VB.CommandButton CommandQuitPP 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton CommandOKPP 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label LabelPGCD 
      Caption         =   "PGCD :"
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
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label LabelPPCM 
      Caption         =   "PPCM :"
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
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2Nombre 
      Caption         =   "2ème nombre "
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
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label LabelN2 
      Caption         =   "N2 :"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label LabelN1 
      Caption         =   "N1 :"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1Nombre 
      Caption         =   "1er nombre "
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
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "FenetrePPCMPGCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calcule_PPCMPGCD()
   On Error GoTo Traite_ErreursPP
   TextPPCM.Text = "Calcul en cours..."
   TextPGCD.Text = "Calcul en cours..."
   Apd# = CDbl(TextN1.Text)
   Bpd# = CDbl(TextN2.Text)
   If Apd# <= 0 Or Bpd# <= 0 Then
   MsgBox "N1 et N2 entiers positifs !", 48
   Exit Sub
   End If
   If Apd# < Bpd# Then
   Cpd# = Bpd#
   Bpd# = Apd#
   Apd# = Cpd#
   End If
   A1pd# = Apd#
   B1pd# = Bpd#
   Do
   Qpd = Int(Apd# / Bpd#)
   rt# = Apd# - Bpd# * Qpd
   If rt# = 0 Then Exit Do
   Apd# = Bpd#
   Bpd# = rt#
   Loop
   TextPPCM.Text = Format(A1pd# * B1pd# / Bpd#)
   TextPGCD.Text = Format(Bpd#)
   Exit Sub
Traite_ErreursPP:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub CommandOKPP_Click()
   Call Calcule_PPCMPGCD
End Sub

Private Sub CommandQuitPP_Click()
   FenetrePPCMPGCD.Hide
End Sub

Private Sub TextN1_Change()
    TextPPCM.Text = ""
    TextPGCD.Text = ""
End Sub

Private Sub TextN2_Change()
    TextPPCM.Text = ""
    TextPGCD.Text = ""
End Sub
