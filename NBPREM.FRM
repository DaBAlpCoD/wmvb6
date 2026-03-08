VERSION 5.00
Begin VB.Form FenetreNbPremiers 
   Caption         =   "Liste des nombres premiers"
   ClientHeight    =   5280
   ClientLeft      =   1110
   ClientTop       =   1485
   ClientWidth     =   7365
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
   ScaleHeight     =   5280
   ScaleWidth      =   7365
   Begin VB.CommandButton CommandDebutNbP 
      Caption         =   "Début"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CommandQuitNbP 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton CommandPgSuiNbP 
      Caption         =   "Page suivante"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FenetreNbPremiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xnl#
Const nl% = 20

Private Sub CalculeNbP()
   i% = 0
   ncp# = xnl#
   FenetreNbPremiers.Cls
   FenetreNbPremiers.PSet (0, 800)
   Do
      Do
         ncp# = ncp# + 1
         If fni(ncp# / 5) = 0 Then Exit Do
         If fni(ncp# / 6) = 0 Then Exit Do
         For f% = 2 To 4
            If fni((ncp# - f%) / 6) = 0 Then Exit Do
         Next f%
         For f% = 5 To Sqr(ncp#) Step 2
            If fni(ncp# / f%) = 0 Then Exit Do
         Next f%
         FenetreNbPremiers.Print ncp#
         i% = i% + 1
         If i% = nl% Then xnl# = ncp#
      Loop Until i% = nl%
    Loop Until i% = nl%
End Sub

Private Sub CommandDebutNbP_Click()
   xnl# = 0
   CalculeNbP
End Sub

Private Sub CommandPgSuiNbP_Click()
   CalculeNbP
End Sub

Private Sub CommandQuitNbP_Click()
   FenetreNbPremiers.Hide
End Sub

Private Function fni(T As Double) As Double
   fni = T - Int(T)
End Function

Private Sub Form_Activate()
   xnl# = 0
   CalculeNbP
End Sub

