VERSION 5.00
Begin VB.Form FenetreRotaPol 
   Caption         =   "Polyèdre : rotations"
   ClientHeight    =   6240
   ClientLeft      =   300
   ClientTop       =   735
   ClientWidth     =   8820
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
   ScaleWidth      =   8820
   Begin VB.CommandButton cmdRotZPos 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   3
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton cmdRotZNeg 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   4
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton cmdRotYPos 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      TabIndex        =   11
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton cmdRotYNeg 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7680
      TabIndex        =   10
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton cmdRotXPos 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   9
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdRotXNeg 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   8
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox pctRepere 
      Height          =   975
      Left            =   6600
      ScaleHeight     =   915
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
      Begin VB.Label lblZ 
         Alignment       =   1  'Right Justify
         Caption         =   "z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape shapeZ 
         Height          =   255
         Left            =   480
         Shape           =   3  'Circle
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblY 
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Line lineYFG 
         X1              =   480
         X2              =   600
         Y1              =   240
         Y2              =   120
      End
      Begin VB.Line lineYFD 
         X1              =   600
         X2              =   720
         Y1              =   120
         Y2              =   240
      End
      Begin VB.Line lineXFB 
         X1              =   1080
         X2              =   1200
         Y1              =   840
         Y2              =   720
      End
      Begin VB.Line lineXFH 
         X1              =   1080
         X2              =   1200
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line lineYA 
         X1              =   600
         X2              =   600
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line lineXA 
         X1              =   600
         X2              =   1200
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.PictureBox pctVuePol 
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblValNbFac 
      Height          =   255
      Left            =   8280
      TabIndex        =   21
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblValNbAr 
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblValNbSom 
      Height          =   255
      Left            =   8280
      TabIndex        =   19
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblNbFaces 
      Caption         =   "Nombre de faces :"
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblNbAr 
      Caption         =   "Nombre d'arêtes :"
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblNbSom 
      Caption         =   "Nombre de sommets :"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblMessage 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line3 
      X1              =   6240
      X2              =   8640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   6240
      X2              =   8640
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblOz 
      Caption         =   "Oz"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblOy 
      Caption         =   "Oy"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblOx 
      Caption         =   "Ox"
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblRotAxe 
      Caption         =   "Rotation  autour  de  l'axe  :"
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
      Left            =   6240
      TabIndex        =   12
      Top             =   2160
      Width           =   2415
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
Attribute VB_Name = "FenetreRotaPol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AnglElem   ' angle de rotation élémentaire
Dim DEP%       ' largeur et hauteur de pctVuePol  (en twips)
Dim DAPol
Dim MPol
Dim RotFlag%

Private Sub AretesVisibles()
' ----------------------------------
' Routine arêtes visibles et cachées
' ----------------------------------
For i% = 1 To NbSom% - 1
   For j% = i% To NbSom%
      If Arete%(i%, j%) = 1 Then Arete%(i%, j%) = 2
   Next j%
Next i%
For jf% = 1 To NbFac%
   X1 = Xsom(NumSomFac%(jf%, 1))
   Y1 = Ysom(NumSomFac%(jf%, 1))
   Z1 = Zsom(NumSomFac%(jf%, 1))
   X2 = Xsom(NumSomFac%(jf%, 2))
   Y2 = Ysom(NumSomFac%(jf%, 2))
   Z2 = Zsom(NumSomFac%(jf%, 2))
   X3 = Xsom(NumSomFac%(jf%, 3))
   Y3 = Ysom(NumSomFac%(jf%, 3))
   Z3 = Zsom(NumSomFac%(jf%, 3))
   ZP = SignePV%(jf%) * ((X2 - X1) * (Y3 - Y1) - (Y2 - Y1) * (X3 - X1))
   If ZP > 0 Then
      NS% = NbSomFac%(jf%)
      For i% = 1 To NS% - 1
         For j% = i% To NS%
            If Arete%(NumSomFac%(jf%, i%), NumSomFac%(jf%, j%)) <> 0 Then
               Arete%(NumSomFac%(jf%, i%), NumSomFac%(jf%, j%)) = 1
            End If
         Next j%
      Next i%
   End If
Next jf%
End Sub


Private Sub cmdRotXNeg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ------ Rotation négative autour de l'axe des X --------
   RotFlag% = 1
   Do
      FenetreRotaPol.pctVuePol.ForeColor = FenetreRotaPol.pctVuePol.BackColor
      Call TraceAretes
      For i% = 1 To NbSom%
         Yloc = Ysom(i%)
         Zloc = Zsom(i%)
         Ysom(i%) = Cos(AnglElem) * Yloc + Sin(AnglElem) * Zloc
         Zsom(i%) = -Sin(AnglElem) * Yloc + Cos(AnglElem) * Zloc
      Next i%
      Call AretesVisibles
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_AVANT
      Call TraceAretes
      DoEvents
   Loop Until RotFlag% = 0
End Sub

Private Sub cmdRotXNeg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RotFlag% = 0
End Sub

Private Sub cmdRotXPos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ------ Rotation positive autour de l'axe des X --------
   RotFlag% = 1
   Do
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_ARRIERE
      Call TraceAretes
      For i% = 1 To NbSom%
         Yloc = Ysom(i%)
         Zloc = Zsom(i%)
         Ysom(i%) = Cos(AnglElem) * Yloc - Sin(AnglElem) * Zloc
         Zsom(i%) = Sin(AnglElem) * Yloc + Cos(AnglElem) * Zloc
      Next i%
      Call AretesVisibles
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_AVANT
      Call TraceAretes
      DoEvents
   Loop Until RotFlag% = 0
End Sub

Private Sub cmdRotXPos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RotFlag% = 0
End Sub

Private Sub cmdRotYNeg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ------ Rotation négative autour de l'axe des Y --------
   RotFlag% = 1
   Do
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_ARRIERE
      Call TraceAretes
      For i% = 1 To NbSom%
         Zloc = Zsom(i%)
         Xloc = Xsom(i%)
         Zsom(i%) = Cos(AnglElem) * Zloc + Sin(AnglElem) * Xloc
         Xsom(i%) = -Sin(AnglElem) * Zloc + Cos(AnglElem) * Xloc
      Next i%
      Call AretesVisibles
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_AVANT
      Call TraceAretes
      DoEvents
   Loop Until RotFlag% = 0
End Sub

Private Sub cmdRotYNeg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RotFlag% = 0
End Sub

Private Sub cmdRotYPos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ------ Rotation positive autour de l'axe des Y --------
   RotFlag% = 1
   Do
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_ARRIERE
      Call TraceAretes
      For i% = 1 To NbSom%
         Zloc = Zsom(i%)
         Xloc = Xsom(i%)
         Zsom(i%) = Cos(AnglElem) * Zloc - Sin(AnglElem) * Xloc
         Xsom(i%) = Sin(AnglElem) * Zloc + Cos(AnglElem) * Xloc
      Next i%
      Call AretesVisibles
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_AVANT
      Call TraceAretes
      DoEvents
   Loop Until RotFlag% = 0
End Sub

Private Sub cmdRotYPos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RotFlag% = 0
End Sub

Private Sub cmdRotZNeg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ------ Rotation négative autour de l'axe des Z --------
   RotFlag% = 1
   Do
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_ARRIERE
      Call TraceAretes
      For i% = 1 To NbSom%
         Xloc = Xsom(i%)
         Yloc = Ysom(i%)
         Xsom(i%) = Cos(AnglElem) * Xloc + Sin(AnglElem) * Yloc
         Ysom(i%) = -Sin(AnglElem) * Xloc + Cos(AnglElem) * Yloc
      Next i%
      Call AretesVisibles
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_AVANT
      Call TraceAretes
      DoEvents
   Loop Until RotFlag% = 0
End Sub

Private Sub cmdRotZNeg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RotFlag% = 0
End Sub

Private Sub cmdRotZPos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ------ Rotation positive autour de l'axe des Z --------
   RotFlag% = 1
   Do
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_ARRIERE
      Call TraceAretes
      For i% = 1 To NbSom%
         Xloc = Xsom(i%)
         Yloc = Ysom(i%)
         Xsom(i%) = Cos(AnglElem) * Xloc - Sin(AnglElem) * Yloc
         Ysom(i%) = Sin(AnglElem) * Xloc + Cos(AnglElem) * Yloc
      Next i%
      Call AretesVisibles
      FenetreRotaPol.pctVuePol.ForeColor = COULEUR_AVANT
      Call TraceAretes
      DoEvents
   Loop Until RotFlag% = 0
End Sub

Private Sub cmdRotZPos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RotFlag% = 0
End Sub

Private Sub Form_Activate()
   FenetreRotaPol.pctVuePol.Cls
   lblValNbSom.Caption = NbSom%
   lblValNbAr.Caption = ""
   lblValNbFac.Caption = ""
   lblMessage.Caption = "Calcul en cours..."
   DoEvents
   Call TrouveAretes
   lblMessage.Caption = NomPolyedre$
   If NomPolyedre$ = "" Then
      ' polyèdre défini par l'utilisateur
      Call AretesVisibles
   End If
   lblMessage.Caption = NomPolyedre$
   lblValNbSom.Caption = CStr(NbSom%)
   lblValNbAr.Caption = CStr(NbAr%)
   lblValNbFac.Caption = CStr(NbFac%)
   '-----------------------------------------------
   ' **********************************************
   ' ******     paramètres pour le tracé     ******
   ' ******  dans la zone d'image pctVuePol  ******
   ' **********************************************
   DEP% = FenetreRotaPol.pctVuePol.ScaleWidth
   DAPol = 1.2 * RayonMax
   MPol = DEP% / DAPol / 2
   '-----------------------------------------------
   Call AretesVisibles
   Call TraceAretes
End Sub

Private Sub Form_Load()
   AnglElem = PI / 1800
End Sub

Private Sub TraceAretes()
' --------------------
' Routine tracé arêtes
' --------------------
For i% = 1 To NbSom% - 1
   For j% = i% + 1 To NbSom%
      If Arete%(i%, j%) <> 0 Then
         If Arete%(i%, j%) = 1 Then
            FenetreRotaPol.pctVuePol.DrawStyle = 0  ' trait plein
         ElseIf Arete%(i%, j%) = 2 Then
            FenetreRotaPol.pctVuePol.DrawStyle = 2  ' tirets
         End If
         pctVuePol.Line (XPdes(Xsom(i%)), YPdes(Ysom(i%)))-(XPdes(Xsom(j%)), YPdes(Ysom(j%)))
      End If
   Next j%
Next i%
End Sub

Private Sub TrouveAretes()
'----- centre de gravité -----
XG = 0
YG = 0
ZG = 0
For i% = 1 To NbSom%
   XG = XG + Xsom(i%)
   YG = YG + Ysom(i%)
   ZG = ZG + Zsom(i%)
Next i%
XG = XG / NbSom%
YG = YG / NbSom%
ZG = ZG / NbSom%
'----- origine au centre de gravité -----
RayonMax = 0
For i% = 1 To NbSom%
Xsom(i%) = Xsom(i%) - XG
Ysom(i%) = Ysom(i%) - YG
Zsom(i%) = Zsom(i%) - ZG
Rayon(i%) = Sqr(Xsom(i%) * Xsom(i%) + Ysom(i%) * Ysom(i%) + Zsom(i%) * Zsom(i%))
If RayonMax < Rayon(i%) Then RayonMax = Rayon(i%)
Next i%
' ----- recherche des faces -----
eps = 0.0001
eps1 = 0.001
NbFac% = 0
NbAr% = 0
jf% = 0
For i% = 1 To NbSom% - 2
   For j% = i% To NbSom% - 1
      For k% = j% To NbSom%
         X1 = Xsom(i%)
         Y1 = Ysom(i%)
         Z1 = Zsom(i%)
         X2 = Xsom(j%)
         Y2 = Ysom(j%)
         Z2 = Zsom(j%)
         X3 = Xsom(k%)
         Y3 = Ysom(k%)
         Z3 = Zsom(k%)
         XV = (Y2 - Y1) * (Z3 - Z1) - (Z2 - Z1) * (Y3 - Y1)
         YV = (Z2 - Z1) * (X3 - X1) - (X2 - X1) * (Z3 - Z1)
         ZV = (X2 - X1) * (Y3 - Y1) - (Y2 - Y1) * (X3 - X1)
         s = XV * X1 + YV * Y1 + ZV * Z1
         If Abs(s) > eps Then
            SE = Sgn(s)
            jf% = jf% + 1
            SignePV%(jf%) = SE
            XV = SE * XV
            YV = SE * YV
            ZV = SE * ZV
            NS% = 3
            For l% = 1 To NbSom%
               ExitFlag1% = 0
               If l% <> i% And l% <> j% And l% <> k% Then
                  p = XV * (X1 - Xsom(l%)) + YV * (Y1 - Ysom(l%)) + ZV * (Z1 - Zsom(l%))
                  If Abs(p) < eps Then
                     NS% = NS% + 1
                     NumSomFac%(jf%, NS%) = l%
                  ElseIf Sgn(p) = -1 Then
                     jf% = jf% - 1
                     ExitFlag1% = 1
                     Exit For
                  End If
               End If
            Next l%
            If ExitFlag1% = 0 Then
               NumSomFac%(jf%, 1) = i%
               NumSomFac%(jf%, 2) = j%
               NumSomFac%(jf%, 3) = k%
               NbSomFac%(jf%) = NS%
               If NS% = 3 Then
                  Arete%(i%, j%) = 2
                  Arete%(i%, k%) = 2
                  Arete%(j%, k%) = 2
                  NbAr% = NbAr% + 3
               Else
                  Rem --- Début: faces polygonales de plus de trois sommets ---
                  Rem Centre de gravité de la face
                  XF = 0
                  YF = 0
                  ZF = 0
                  For f% = 1 To NS%
                  XF = XF + Xsom(NumSomFac%(jf%, f%))
                  YF = YF + Ysom(NumSomFac%(jf%, f%))
                  ZF = ZF + Zsom(NumSomFac%(jf%, f%))
                  Next f%
                  XF = XF / NS%
                  Xfac(jf%) = XF
                  YF = YF / NS%
                  Yfac(jf%) = YF
                  ZF = ZF / NS%
                  Zfac(jf%) = ZF
                  KF% = 0
                  Do
                     KF% = KF% + 1
                     If KF% >= jf% Then Exit Do
                     ExitFlag2% = 0
                     If Abs(XF - Xfac(KF%)) < eps1 And Abs(YF - Yfac(KF%)) < eps1 And Abs(ZF - Zfac(KF%)) < eps1 Then
                        jf% = jf% - 1
                        ExitFlag2% = 1
                        Exit Do
                     End If
                  Loop
                  If ExitFlag2% = 0 Then
                     For f% = 1 To NS% - 1
                        For g% = f% To NS%
                           X1 = Xsom(NumSomFac%(jf%, f%))
                           Y1 = Ysom(NumSomFac%(jf%, f%))
                           Z1 = Zsom(NumSomFac%(jf%, f%))
                           X2 = Xsom(NumSomFac%(jf%, g%))
                           Y2 = Ysom(NumSomFac%(jf%, g%))
                           Z2 = Zsom(NumSomFac%(jf%, g%))
                           XW = (Y2 - Y1) * (ZF - Z1) - (Z2 - Z1) * (YF - Y1)
                           YW = (Z2 - Z1) * (XF - X1) - (X2 - X1) * (ZF - Z1)
                           ZW = (X2 - X1) * (YF - Y1) - (Y2 - Y1) * (XF - X1)
                           lv = XW * XW + YW * YW + ZW * ZW
                           If lv > eps Then
                              For h% = 1 To NS%
                                 ExitFlag3% = 0
                                 If h% <> f% And h% <> g% Then
                                    X3 = Xsom(NumSomFac%(jf%, h%))
                                    Y3 = Ysom(NumSomFac%(jf%, h%))
                                    Z3 = Zsom(NumSomFac%(jf%, h%))
                                    XH = (Y2 - Y1) * (Z3 - Z1) - (Z2 - Z1) * (Y3 - Y1)
                                    YH = (Z2 - Z1) * (X3 - X1) - (X2 - X1) * (Z3 - Z1)
                                    ZH = (X2 - X1) * (Y3 - Y1) - (Y2 - Y1) * (X3 - X1)
                                    p = XW * XH + YW * YH + ZW * ZH
                                    If p < 0 Then
                                       ExitFlag3% = 1
                                       Exit For
                                    End If
                                 End If
                              Next h%
                              If ExitFlag3% = 0 Then
                                 Arete%(NumSomFac%(jf%, f%), NumSomFac%(jf%, g%)) = 2
                                 NbAr% = NbAr% + 1
                              End If
                           End If
                        Next g%
                     Next f%
                  End If
                  Rem --- Fin: faces polygonales de plus de trois sommets ---
               End If
            End If
         End If
      Next k%
   Next j%
Next i%
NbFac% = jf%
NbAr% = NbAr% / 2
End Sub

Private Function XPdes(Xloc)
   XPdes = (DAPol + Xloc) * MPol
End Function

Private Function YPdes(Yloc)
   YPdes = (DAPol - Yloc) * MPol
End Function



Private Sub mnuQuitter_Click()
   FenetreRotaPol.Hide
   FenetrePolyedre.Hide
End Sub

Private Sub mnuRetour_Click()
   FenetreRotaPol.Hide
End Sub


