VERSION 5.00
Begin VB.Form FenetreVisuTriangle 
   Caption         =   "Triangle : visualiser"
   ClientHeight    =   6240
   ClientLeft      =   300
   ClientTop       =   735
   ClientWidth     =   8670
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
   ScaleHeight     =   6240
   ScaleWidth      =   8670
   Begin VB.CommandButton cmdECercEul 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdECerCir 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdECercIns 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdEMediat 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdTCercEul 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdTCerCir 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdTCercIns 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdTMediat 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdTBissec 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEBissec 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdEHauteur 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdTHauteur 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdEMedian 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdTMedian 
      Caption         =   "Tracer"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox pctVueTrian 
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblCerCir 
      Caption         =   "Cercle circonscrit :"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblCercIns 
      Caption         =   "Cercle inscrit :"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblCercEul 
      Caption         =   "Cercle d'Euler :"
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblMediat 
      Caption         =   "Médiatrices :"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblBissec 
      Caption         =   "Bissectrices :"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblHauteurs 
      Caption         =   "Hauteurs :"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblMedianes 
      Caption         =   "Médianes :"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
   End
   Begin VB.Menu mnuRetour 
      Caption         =   "&Retour"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreVisuTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Det%     ' largeur et hauteur de pctVueTrian (en twips)
Dim DATrian
Dim MTrian
Dim XP(1 To 3)
Dim YP(1 To 3)
Dim DXP(1 To 3)
Dim DYP(1 To 3)
Dim XMP(1 To 3)
Dim YMP(1 To 3)
Dim CP(1 To 3)
Dim UP(1 To 3)

Private Sub CalculPourTrace()
' ------------------- Calculs pour tracé  --------------------
' Centrage sur cercle circonscrit sur ABC (ou HBC si angle A obtus )
   If Obtus% = 0 Then
      RayonMaxTrian = RayonCirc
      Xcentrecran = Xcir
      Ycentrecran = Ycir
   Else
      i% = Obtus%
      j% = i% + 1
      If j% = 4 Then j% = 1
      k% = j% + 1
      If k% = 4 Then k% = 1
      XP(i%) = Xort
      YP(i%) = Yort
      XP(j%) = Xtrian(j%)
      YP(j%) = Ytrian(j%)
      XP(k%) = Xtrian(k%)
      YP(k%) = Ytrian(k%)
      For i% = 1 To 3
         j% = i% + 1
         If j% = 4 Then j% = 1
         k% = j% + 1
         If k% = 4 Then k% = 1
         DXP(i%) = XP(k%) - XP(j%)
         DYP(i%) = YP(k%) - YP(j%)
         CP(i%) = Sqr(DXP(i%) * DXP(i%) + DYP(i%) * DYP(i%))
         XMP(i%) = (XP(j%) + XP(k%)) / 2
         YMP(i%) = (YP(j%) + YP(k%)) / 2
      Next i%
      LP = DXP(1) * DYP(2) - DXP(2) * DYP(1)
      Xcentrecran = (2 * XMP(1) * DXP(1) * DYP(2) + DYP(1) * DYP(2) * DYP(3) - 2 * XMP(2) * DXP(2) * DYP(1)) / 2 / LP
      Ycentrecran = -(2 * YMP(1) * DYP(1) * DXP(2) + DXP(1) * DXP(2) * DXP(3) - 2 * YMP(2) * DYP(2) * DXP(1)) / 2 / LP
      PP = CP(1) + CP(2) + CP(3): PP = PP / 2
      SP = 0
      RayonMaxTrian = 1
      For i% = 1 To 3
      UP(i%) = PP - CP(i%)
      SP = SP - XP(i%) * DYP(i%)
      RayonMaxTrian = RayonMaxTrian * CP(i%)
      Next i%
      SP = SP / 2
      RayonMaxTrian = RayonMaxTrian / SP / 4
   End If
End Sub

Private Sub cmdEBissec_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))-(XTdes(Xbissec(i%)), YTdes(Ybissec(i%)))
   Next i%
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub

Private Sub cmdECercEul_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   pctVueTrian.Circle (XTdes(Xeul), YTdes(Yeul)), RTdes(RayonEuler)
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub

Private Sub cmdECercIns_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   pctVueTrian.Circle (XTdes(Xins), YTdes(Yins)), RTdes(RayonInscrit)
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub

Private Sub cmdECerCir_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   pctVueTrian.Circle (XTdes(Xcir), YTdes(Ycir)), RTdes(RayonCirc)
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub

Private Sub cmdEHauteur_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))-(XTdes(Xhauteur(i%)), YTdes(Yhauteur(i%)))
      If Obtus% <> 0 Then
         pctVueTrian.DrawStyle = 2  ' pointillés
         pctVueTrian.Line (XTdes(Xort), YTdes(Yort))-(XTdes(Xhauteur(i%)), YTdes(Yhauteur(i%)))
         pctVueTrian.DrawStyle = 0  ' continu
      End If
   Next i%
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub

Private Sub cmdEMedian_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))-(XTdes(Xmediane(i%)), YTdes(Ymediane(i%)))
   Next i%
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub

Private Sub cmdEMediat_Click()
   pctVueTrian.ForeColor = pctVueTrian.BackColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xmediat(i%)), YTdes(Ymediat(i%)))-(XTdes(Xmediane(i%)), YTdes(Ymediane(i%)))
      If Obtus% <> 0 Then
         pctVueTrian.DrawStyle = 2  ' pointillés
         pctVueTrian.Line (XTdes(Xcir), YTdes(Ycir))-(XTdes(Xmediat(i%)), YTdes(Ymediat(i%)))
         pctVueTrian.DrawStyle = 0  ' continu
      End If
   Next i%
   '---------------------------------------------------
   ' Retracer le triangle
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
   '---------------------------------------------------
End Sub


Private Sub cmdTBissec_Click()
   pctVueTrian.ForeColor = lblBissec.ForeColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))-(XTdes(Xbissec(i%)), YTdes(Ybissec(i%)))
   Next i%
End Sub

Private Sub cmdTCercEul_Click()
   pctVueTrian.ForeColor = lblCercEul.ForeColor
   pctVueTrian.Circle (XTdes(Xeul), YTdes(Yeul)), RTdes(RayonEuler)
End Sub

Private Sub cmdTCercIns_Click()
   pctVueTrian.ForeColor = lblCercIns.ForeColor
   pctVueTrian.Circle (XTdes(Xins), YTdes(Yins)), RTdes(RayonInscrit)
End Sub

Private Sub cmdTCerCir_Click()
   pctVueTrian.ForeColor = lblCerCir.ForeColor
   pctVueTrian.Circle (XTdes(Xcir), YTdes(Ycir)), RTdes(RayonCirc)
End Sub

Private Sub cmdTHauteur_Click()
   pctVueTrian.ForeColor = lblHauteurs.ForeColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))-(XTdes(Xhauteur(i%)), YTdes(Yhauteur(i%)))
      If Obtus% <> 0 Then
         pctVueTrian.DrawStyle = 2  ' pointillés
         pctVueTrian.Line (XTdes(Xort), YTdes(Yort))-(XTdes(Xhauteur(i%)), YTdes(Yhauteur(i%)))
         pctVueTrian.DrawStyle = 0  ' continu
      End If
   Next i%
End Sub

Private Sub cmdTMedian_Click()
   pctVueTrian.ForeColor = lblMedianes.ForeColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))-(XTdes(Xmediane(i%)), YTdes(Ymediane(i%)))
   Next i%
End Sub

Private Sub cmdTMediat_Click()
   pctVueTrian.ForeColor = lblMediat.ForeColor
   For i% = 1 To 3
      pctVueTrian.Line (XTdes(Xmediat(i%)), YTdes(Ymediat(i%)))-(XTdes(Xmediane(i%)), YTdes(Ymediane(i%)))
      If Obtus% <> 0 Then
         pctVueTrian.DrawStyle = 2  ' pointillés
         pctVueTrian.Line (XTdes(Xcir), YTdes(Ycir))-(XTdes(Xmediat(i%)), YTdes(Ymediat(i%)))
         pctVueTrian.DrawStyle = 0  ' continu
      End If
   Next i%
End Sub

Private Sub Form_Activate()
   Call CalculPourTrace
   '--------------------------------------------------
   ' *************************************************
   ' ****       paramètres pour le tracé        ******
   ' ****  dans la zone d'image pctVueTriangle  ******
   ' *************************************************
   Det% = FenetreVisuTriangle.pctVueTrian.ScaleWidth
   DATrian = 1.2 * RayonMaxTrian
   MTrian = Det% / DATrian / 2
   '---------------------------------------------------
   FenetreVisuTriangle.pctVueTrian.Cls
   FenetreVisuTriangle.pctVueTrian.ForeColor = NOIR
   FenetreVisuTriangle.pctVueTrian.PSet (XTdes(Xtrian(3)), YTdes(Ytrian(3)))
   For i% = 1 To 3
      FenetreVisuTriangle.pctVueTrian.Line -(XTdes(Xtrian(i%)), YTdes(Ytrian(i%)))
   Next i%
End Sub

Private Sub Form_Load()
   lblMedianes.ForeColor = BLEU
   lblHauteurs.ForeColor = ROUGE
   lblBissec.ForeColor = CYAN
   lblMediat.ForeColor = MAGENTA
   lblCercIns.ForeColor = lblBissec.ForeColor
   lblCerCir.ForeColor = lblMediat.ForeColor
   lblCercEul.ForeColor = VERT
End Sub

Private Function RTdes(Rloc)
   RTdes = Rloc * MTrian
End Function

Private Function XTdes(Xloc)
   XTdes = (DATrian + Xloc - Xcentrecran) * MTrian
End Function

Private Function YTdes(Yloc)
   YTdes = (DATrian - Yloc + Ycentrecran) * MTrian
End Function



Private Sub mnuQuitter_Click()
   FenetreVisuTriangle.Hide
   FenetreTriangle.Hide
End Sub

Private Sub mnuRetour_Click()
   FenetreVisuTriangle.Hide
End Sub


