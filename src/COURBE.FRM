VERSION 5.00
Begin VB.Form Fenetre_Prin 
   Caption         =   "Courbe "
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11655
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   121.762
   ScaleMode       =   0  'User
   ScaleWidth      =   11655
   Begin VB.Label LabelExplications 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
         Begin VB.Menu mnuOuvImage 
            Caption         =   "Une &image..."
         End
         Begin VB.Menu mnuOuvPoints 
            Caption         =   "Un &ensemble de points..."
         End
      End
      Begin VB.Menu mnuEnregistrer 
         Caption         =   "En&registrer..."
         Begin VB.Menu mnuEnrImage 
            Caption         =   "Une i&mage..."
         End
         Begin VB.Menu mnuEnrPoints 
            Caption         =   "Un e&nsemble de points..."
         End
      End
   End
   Begin VB.Menu mnuTraceStop 
      Caption         =   "&Tracer"
   End
   Begin VB.Menu mnuImprimer 
      Caption         =   "Im&primer"
   End
   Begin VB.Menu mnuRetour 
      Caption         =   "&Retour"
   End
   Begin VB.Menu mnuAide 
      Caption         =   "&?"
      Begin VB.Menu mnuComProc 
         Caption         =   "&Comment procéder"
      End
   End
End
Attribute VB_Name = "Fenetre_Prin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YA



Private Sub EulerED2Moins()
'----------------------------------------
' Méthode d'EULER
' Equation différentielle du second ordre
' (X décroissant):
' Y'i+1 = Y'i + F(Xi,Yi,Y'i) dX
' Yi+1 = Yi + Y'i dX
' On pose Y'=Z
'----------------------------------------
' Progression vers les X négatifs :
' dx < 0    (dx = -RX)
'----------------------------------------
'------- Calcul de Yi+1 et Y'i+1 : ------
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
vvar$(3) = LTrim$(Format(Z, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
Y = Y - Z * RX
Z = Z - f * RX
'----------- incrémentation de X --------
X = X - RX
'----------------------------------------
End Sub

Private Sub EulerED2Plus()
'----------------------------------------
' Méthode d'EULER
' Equation différentielle du second ordre
' (X croissant):
' Y'i+1 = Y'i + F(Xi,Yi,Y'i) dX
' Yi+1 = Yi + Y'i dX
' On pose Y'=Z
'----------------------------------------
'------- Calcul de Y'i+1 et Y'i+1 : -----
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
vvar$(3) = LTrim$(Format(Z, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
Y = Y + Z * RX
Z = Z + f * RX
'----------- incrémentation de X --------
X = X + RX
'----------------------------------------
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
    Fenetre_Val.ForeColor = Fenetre_Val.BackColor
    Fenetre_Val.PSet (0, 3900), Fenetre_Val.BackColor
    Fenetre_Val.Print Space$(50)
    Fenetre_Val.PSet (0, 4200), Fenetre_Val.BackColor
    Fenetre_Val.Print Space$(50)
    Fenetre_Val.Print Space$(50)
End Sub


Private Sub GraduAx()
    Rem ----- Routine graduation cadre et/ou axes et grille -----
    LG% = Int((LX% + LY%) / 2)
    Rem <<< détermination de l'intervalle de graduation >>>
    LAY = M * Log(AY)
    ILAY% = Int(LAY)
    YP = 10 ^ ILAY%
    Do
       YP = YP / 2
       If AY / YP > 16 Then Exit Do
       YP = YP / 5
    Loop Until AY / YP > 16
    VY = 10 * YP
    YP = VY
    Do
       YP = YP / 10
    Loop Until AY / YP > 40
    GY = 10 * YP
    GY = 10 ^ CInt(M * Log(GY))
    If Ech% = 0 Then
       LAX = M * Log(AX)
       ILAX% = Int(LAX)
       XP = 10 ^ ILAX%
       Do
          XP = XP / 2
          If AX / XP > 24 Then Exit Do
          XP = XP / 5
       Loop Until AX / XP > 24
       VX = 10 * XP
       XP = VX
       Do
          XP = XP / 10
       Loop Until AX / XP > 60
       GX = 10 * XP
       GX = 10 ^ CInt(M * Log(GX))
    ElseIf Ech% = 1 Then
       GX = GY
       VX = VY
    End If
    Rem <<< graduation en x >>>
    X = GX * Int(Xmin / GX + 1)
    Do While X < Xmax
       If Grax% = 1 Then
          Y = 0
          Dessin.Line (Xdes(X), Ydes(Y) - 0.01 * LG%)-(Xdes(X), Ydes(Y) + 0.01 * LG%)
          If Vax% = 1 Then
             If Abs(X / VX - CInt(X / VX)) < GX / 20 Then
                Dessin.Line (Xdes(X), Ydes(Y) - 0.015 * LG%)-(Xdes(X), Ydes(Y) + 0.015 * LG%)
                X0 = VX * CInt(X / VX)
                Call ValGradX(X0, Y)
             End If
          End If
       End If
       If Tric% = 1 Then
          Dessin.DrawStyle = 2
          Dessin.Line (Xdes(X), Ydes(Ymax))-(Xdes(X), Ydes(Ymin))
          Dessin.DrawStyle = 0
       End If
       If Grac% = 1 Then
          Y = Ymin
          Dessin.Line (Xdes(X), Ydes(Y) - 0.01 * LG%)-(Xdes(X), Ydes(Y) + 0.01 * LG%)
          Y = Ymax
          Dessin.Line (Xdes(X), Ydes(Y) - 0.01 * LG%)-(Xdes(X), Ydes(Y) + 0.01 * LG%)
          If Vac% = 1 Then
             If Abs(X / VX - CInt(X / VX)) < GX / 20 Then
                Dessin.Line (Xdes(X), Ydes(Y) - 0.015 * LG%)-(Xdes(X), Ydes(Y) + 0.015 * LG%)
                X0 = VX * CInt(X / VX)
                Call ValGradX(X0, Y)
             End If
          End If
       End If
       X = X + GX
    Loop
    If Grax% = 1 Then
       Dessin.Line (Xdes(Xmax), Ydes(0))-(Xdes(Xmax) - 0.02 * LG%, Ydes(0) + 0.02 * LG%)
       Dessin.Line (Xdes(Xmax), Ydes(0))-(Xdes(Xmax) - 0.02 * LG%, Ydes(0) - 0.02 * LG%)
    End If
    Rem <<< graduation en y >>>
    Y = GY * Int(Ymin / GY + 1)
    Do While Y < Ymax
       If Grax% = 1 Then
          X = 0
          Dessin.Line (Xdes(X) - 0.01 * LG%, Ydes(Y))-(Xdes(X) + 0.01 * LG%, Ydes(Y))
          If Vax% = 1 Then
             If Abs(Y / VY - CInt(Y / VY)) < GY / 20 Then
                Dessin.Line (Xdes(X) - 0.015 * LG%, Ydes(Y))-(Xdes(X) + 0.015 * LG%, Ydes(Y))
                Y0 = VY * CInt(Y / VY)
                If Y0 <> 0 Then Call ValGradY(X, Y0)
             End If
          End If
       End If
       If Tric% = 1 Then
          Dessin.DrawStyle = 2
          Dessin.Line (Xdes(Xmax), Ydes(Y))-(Xdes(Xmin), Ydes(Y))
          Dessin.DrawStyle = 0
       End If
       If Grac% = 1 Then
          X = Xmax
          Dessin.Line (Xdes(X) - 0.01 * LG%, Ydes(Y))-(Xdes(X) + 0.01 * LG%, Ydes(Y))
          X = Xmin
          Dessin.Line (Xdes(X) - 0.01 * LG%, Ydes(Y))-(Xdes(X) + 0.01 * LG%, Ydes(Y))
          If Vac% = 1 Then
             If Abs(Y / VY - CInt(Y / VY)) < GY / 20 Then
                Dessin.Line (Xdes(X) - 0.015 * LG%, Ydes(Y))-(Xdes(X) + 0.015 * LG%, Ydes(Y))
                Y0 = VY * CInt(Y / VY)
                Call ValGradY(X, Y0)
             End If
          End If
       End If
       Y = Y + GY
    Loop
    If Grax% = 1 Then
       Dessin.Line (Xdes(0), Ydes(Ymax))-(Xdes(0) - 0.02 * LG%, Ydes(Ymax) + 0.02 * LG%)
       Dessin.Line (Xdes(0), Ydes(Ymax))-(Xdes(0) + 0.02 * LG%, Ydes(Ymax) + 0.02 * LG%)
    End If
End Sub



Private Sub mnuComProc_Click()
    Aide.Show
End Sub

Private Sub mnuDefCou_Click()
   Select Case NumChoisi%
   Case NUMCOUR
      FenetreDefCourbe.Show
   Case NUMSOLEQ
      FenetreDefSolEq.Show
   Case NUMEQUADIF
      FenetreDefEqDif.Show
   Case NUMEQUADIF2
      FenetreDefEqDif2.Show
   Case NUMPRIMITIVE
      FenetreDefPrim.Show
   Case NUMINTEGRALE
      FenetreDefInteg.Show
   'Case NUMOINCAR
   '   FenetreDefMoinCar.Show
   'Case NUMREGRELIN
   '   FenetreDefRegreLin.Show
   'Case NUMINPOLAG
   '   FenetreDefInPoLag.Show
   'Case NUMSPLINE
   '   FenetreDefSpline.Show
   Case NUMPOLYNOME
      FenetreDefPolynome.Show
   End Select
End Sub

Private Sub mnuDefinir_Click()
   Select Case NumChoisi%
   Case NUMCOUR, NUMSOLEQ
   '   FenetreDefCourbe.Show
   Case NUMEQUADIF
   '   FenetreDefEqDif.Show
   Case NUMEQUADIF2
   '   FenetreDefEqDif2.Show
   Case NUMPRIMITIVE
   '   FenetreDefPrim.Show
   Case NUMINTEGRALE
   '   FenetreDefInteg.Show
   Case NUMOINCAR
      FenetreDefMoinCar.Show
   Case NUMREGRELIN
      FenetreDefRegreLin.Show
   Case NUMINPOLAG
      FenetreDefInPoLag.Show
   Case NUMSPLINE
      FenetreDefSpline.Show
   Case NUMPOLYNOME
   '   FenetreDefPolynome.Show
   End Select
End Sub




Private Sub mnuEnrImage_Click()
   On Error GoTo Traite_ErreursEnrIm
   Maths.ctrlCMDialog.DefaultExt = "bmp"
   Maths.ctrlCMDialog.Filter = "Image (*.bmp)|*.bmp"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   SavePicture Dessin.Image, Maths.ctrlCMDialog.FileName
   Exit Sub
Traite_ErreursEnrIm:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuEnrPoints_Click()
   On Error GoTo Traite_ErreursEnrpts
   Maths.ctrlCMDialog.DefaultExt = "pts"
   Maths.ctrlCMDialog.Filter = "Points (*.xy)|*.xy"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   '-----------------------------------------------------
   ' Création du fichier de points
   ' et écriture de leurs coordonnées Xpt() et Ypt()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Output As #1
   Select Case NumChoisi%
   Case NUMCOUR, NUMSOLEQ, NUMEQUADIF, NUMEQUADIF2, NUMPRIMITIVE, NUMINTEGRALE, NUMPOLYNOME
      Write #1, NbPts%
      For iloc% = 1 To NbPts%
         Write #1, Xpt(iloc%)
         Write #1, Ypt(iloc%)
      Next iloc%
   Case NUMOINCAR
      Write #1, NbPtsMoinCar%
      For iloc% = 1 To NbPtsMoinCar%
         Write #1, Xmc(iloc%)
         Write #1, Ymc(iloc%)
      Next iloc%
   Case NUMREGRELIN
      Write #1, NbPtsReLin%
      For iloc% = 1 To NbPtsReLin%
         Write #1, Xrl(iloc%)
         Write #1, Yrl(iloc%)
      Next iloc%
   Case NUMINPOLAG
      Write #1, NbPtsI%
      For iloc% = 1 To NbPtsI%
         Write #1, Xpl(iloc%)
         Write #1, Ypl(iloc%)
      Next iloc%
   Case NUMSPLINE
      Write #1, NbPtsSp%
      For iloc% = 1 To NbPtsSp%
         Write #1, Xsp(iloc%)
         Write #1, Ysp(iloc%)
      Next iloc%
   End Select
   Close #1
   '-----------------------------------------------------
   Exit Sub
Traite_ErreursEnrpts:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuImprimer_Click()
   On Error GoTo Traite_ErreursImp
   'On Local Error Resume Next
   Maths.ctrlCMDialog.Flags = &H8&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 5
   Dessin.PrintForm
   Exit Sub
Traite_ErreursImp:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuOuvImage_Click()
   On Error GoTo Traite_ErreursOuvIm
   Maths.ctrlCMDialog.Filter = "Image (*.bmp)|*.bmp"
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   Dessin.Picture = LoadPicture(Maths.ctrlCMDialog.FileName)
   Fenetre_Val.lblFormule.Caption = ""
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_ErreursOuvIm:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuOuvPoints_Click()
   On Error GoTo Traite_ErreursOuvPts
   Maths.ctrlCMDialog.Filter = "Points (*.xy)|*.xy"
   ' nom de fichier et chemin doivent exister
   ' sinon apparait un message d'erreur spécifique
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   '-----------------------------------------------------
   ' Ouverture et lecture du fichier de points
   ' et écriture de leurs coordonnées dans Xpt() et Ypt()
   ' ou autres...
   '-----------------------------------------------------
   Open Maths.ctrlCMDialog.FileName For Input As #1
   Select Case NumChoisi%
   Case NUMCOUR, NUMSOLEQ, NUMEQUADIF, NUMEQUADIF2, NUMPRIMITIVE, NUMINTEGRALE, NUMPOLYNOME
      Input #1, NbPts%
      ReDim Xpt(1 To NbPts%)
      ReDim Ypt(1 To NbPts%)
      For iloc% = 1 To NbPts%
         Input #1, Xpt(iloc%)
         Input #1, Ypt(iloc%)
      Next iloc%
   Case NUMOINCAR
      Input #1, NbPtsMoinCar%
      ReDim Xmc(1 To NbPtsMoinCar%)
      ReDim Ymc(1 To NbPtsMoinCar%)
      For iloc% = 1 To NbPtsMoinCar%
         Input #1, Xmc(iloc%)
         Input #1, Ymc(iloc%)
      Next iloc%
   Case NUMREGRELIN
      Input #1, NbPtsReLin%
      ReDim Xrl(1 To NbPtsReLin%)
      ReDim Yrl(1 To NbPtsReLin%)
      For iloc% = 1 To NbPtsReLin%
         Input #1, Xrl(iloc%)
         Input #1, Yrl(iloc%)
      Next iloc%
   Case NUMINPOLAG
      Input #1, NbPtsI%
      DegPoLag% = NbPtsI% - 1
      ReDim Xpl(1 To NbPtsI%)
      ReDim Ypl(1 To NbPtsI%)
      ReDim UparI(0 To NbPtsI% - 1)
      For iloc% = 1 To NbPtsI%
         Input #1, Xpl(iloc%)
         Input #1, Ypl(iloc%)
      Next iloc%
   Case NUMSPLINE
      Input #1, NbPtsSp%
      Erase Xsp, Ysp
      ReDim Xsp(1 To NbPtsSp%)
      ReDim Ysp(1 To NbPtsSp%)
      ReDim Aspline(1 To NbPtsSp% - 1)
      ReDim Bspline(1 To NbPtsSp% - 1)
      ReDim Cspline(1 To NbPtsSp% - 1)
      ReDim Dspline(1 To NbPtsSp% - 1)
      For iloc% = 1 To NbPtsSp%
         Input #1, Xsp(iloc%)
         Input #1, Ysp(iloc%)
      Next iloc%
   End Select
   Close #1
   '-----------------------------------------------------
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_ErreursOuvPts:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         Close #1
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub


Private Sub mnuQuitter_Click()
   Fenetre_Prin.Hide
   Fenetre_Val.Hide
   Dessin.Hide
End Sub


Private Sub RungeKuttaMoins()
'--------------------------------------
' Méthode de RUNGE-KUTTA (X décroissant):
' Yi+1 = Yi + (k1 + 2 k2 + 2 k3 + k4)/6
' avec :
' k1 = F(Xi,Yi) dX
' k2 = F(Xi+dX/2,Yi+k1/2) dX
' k3 = F(Xi+dX/2,Yi+k2/2) dX
' k4 = F(Xi+dX,Yi+k3) dX
'----------------------------------
' Progression vers les X négatifs :
' dx < 0    (dx = -RX)
'----------------------------------
YA = Y
'--------- Calcul de k1 : ---------
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k1 = -f * RX
'--------- Calcul de k2 : ---------
X = X - RX / 2
Y = YA + k1 / 2
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k2 = -f * RX
'--------- Calcul de k3 : ---------
Y = YA + k2 / 2
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k3 = -f * RX
'--------- Calcul de k4 : ---------
X = X - RX / 2
Y = YA + k3
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k4 = -f * RX
'--------- Calcul de Y : ---------
Y = YA + (k1 + 2 * k2 + 2 * k3 + k4) / 6
End Sub

Private Sub RungeKuttaPlus()
'--------------------------------------
' Méthode de RUNGE-KUTTA (X croissant):
' Yi+1 = Yi + (k1 + 2 k2 + 2 k3 + k4)/6
' avec :
' k1 = F(Xi,Yi) dX
' k2 = F(Xi+dX/2,Yi+k1/2) dX
' k3 = F(Xi+dX/2,Yi+k2/2) dX
' k4 = F(Xi+dX,Yi+k3) dX
'----------------------------------
YA = Y
'--------- Calcul de k1 : ---------
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k1 = f * RX
'--------- Calcul de k2 : ---------
X = X + RX / 2
Y = YA + k1 / 2
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k2 = f * RX
'--------- Calcul de k3 : ---------
Y = YA + k2 / 2
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k3 = f * RX
'--------- Calcul de k4 : ---------
X = X + RX / 2
Y = YA + k3
vvar$(1) = LTrim$(Format(X, "0.000000"))
vvar$(2) = LTrim$(Format(Y, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
f = CSng(SorVar2$)
k4 = f * RX
'--------- Calcul de Y : ---------
Y = YA + (k1 + 2 * k2 + 2 * k3 + k4) / 6
End Sub

Private Sub SimpsonPrimMoins()
'--------------------------------------
' Méthode de SIMPSON :
' Aire = (y1 + 4 y2 + 2 y3 + 4 y4 +...+ 4 yi-1 + yi) dX/3
' [ où yi = f(Xi) ]
' Application au calcul de la primitive de f(X) :
' Yi+2 = Yi + (yi + 4 yi+1 + yi+2) dX/3
' en sautant de 2 en 2 dX
'----------------------------------
' Progression vers les X négatifs :
' dx < 0    (dx = -RX/2)
'-------------------------------------------
'--------- Calcul et ajout de yi : ---------
dy = YA
'--------- Calcul et ajout de 4 yi+1 : -----
X = X - RX / 2
vvar$(1) = LTrim$(Format(X, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
YA = CSng(SorVar2$)
dy = dy + 4 * YA
'--------- Calcul et ajout de yi+2 : -------
X = X - RX / 2
vvar$(1) = LTrim$(Format(X, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
YA = CSng(SorVar2$)
dy = dy + YA
'--------- Calcul final : ------------------
dy = dy * -RX / 6
Y = Y + dy
End Sub

Private Sub SimpsonPrimPlus()
'--------------------------------------
' Méthode de SIMPSON :
' Aire = (y1 + 4 y2 + 2 y3 + 4 y4 +...+ 4 yi-1 + yi) dX/3
' [ où yi = f(Xi) ]
' Application au calcul de la primitive de f(X) :
' Yi+2 = Yi + (yi + 4 yi+1 + yi+2) dX/3
' en sautant de 2 en 2 dX
' On prend : dX = RX/2
'-------------------------------------------
'--------- Calcul et ajout de yi : ---------
dy = YA
'--------- Calcul et ajout de 4 yi+1 : -----
X = X + RX / 2
vvar$(1) = LTrim$(Format(X, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
YA = CSng(SorVar2$)
dy = dy + 4 * YA
'--------- Calcul et ajout de yi+2 : -------
X = X + RX / 2
vvar$(1) = LTrim$(Format(X, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
YA = CSng(SorVar2$)
dy = dy + YA
'--------- Calcul final : ------------------
dy = dy * RX / 6
Y = Y + dy
End Sub

Private Sub TraceCou0()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   Fenetre_Val.Cls
   mnuAide.Enabled = False
   mnuDefinir.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   Ymax = FE * (Xmax - Xmin) + Ymin
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If AX < 0 Then
       MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
       Exit Sub
   End If
   AY = AX * FE
   Ymax = Ymin + AY
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '---------------------------------------------------------------
   CouleurCou& = FenetreDefCourbe.pctColorCourbe.BackColor
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
       Dessin.Picture = LoadPicture()
   End If
   On Error GoTo Traite_Erreurs0
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
    If Trac% = 1 Then
        Dessin.Line (0, 0)-(0, LY%)
        Dessin.Line -(LX%, LY%)
        Dessin.Line -(LX%, 0)
        Dessin.Line -(0, 0)
    End If
    If Tric% = 1 Then
        Dessin.Line (0, 0)-(0, LY%)
        Dessin.Line -(LX%, LY%)
        Dessin.Line -(LX%, 0)
        Dessin.Line -(0, 0)
    End If
    If Trax% = 1 Then
        Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
        Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
    End If
    '--------------------------------------------------------------
    '--------- Graduation et grille -------------------------------
    '--------------------------------------------------------------
    If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
    '--------------------------------------------------------------
    '--------- Affichage de la (des) formule(s) -------------------
    '---------     définissant la courbe        -------------------
    '--------------------------------------------------------------
    Fenetre_Val.PSet (0, 0), Fenetre_Val.BackColor
    Fenetre_Val.lblFormule.ForeColor = CouleurCou&
    Fenetre_Val.lblFormule.Caption = " Y = ?"
    '--------------------------------------------------------------
    '--------- Tracé de la courbe ---------------------------------
    '--------------------------------------------------------------
    Dessin.ForeColor = CouleurCou&
    X = Xmin - RX
    SD% = 1
    SE% = 0
    Fenetre_Val.ForeColor = NOIR
    Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
    Fenetre_Val.Print " Un instant SVP ..."
    Do
        DoEvents
        If DoFlag = False Then Exit Do
        Fenetre_Val.ForeColor = Fenetre_Val.BackColor
        Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
        Fenetre_Val.Print Space$(50)
        Fenetre_Val.ForeColor = NOIR
        Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
        Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
        X = X + RX
        If X > Xmax Then Exit Do
        Y = 0.02 * X ^ 6 - 0.02 * X ^ 5 - 0.4 * X ^ 4 + 0.37 * X ^ 3 + 2.18 * X ^ 2 - 1.48 * X - 1.56
        If Erreur = True Then Exit Do
        'SE%=0 : point dans le cadre
        'SE%=1 : point impossible à calculer, erreur
        'SE%=2 : point calculé mais hors cadre
        If SE% <> 1 Then
            If Y > Ymax Or Y < Ymin Then
                SE% = 2
                Select Case SD%
                    Case 0
                        Dessin.Line -(Xdes(X), Ydes(Y))
                    Case 1, 2
                        Dessin.PSet (Xdes(X), Ydes(Y))
                End Select
            Else
                Select Case SD%
                    Case 0, 2
                        Dessin.Line -(Xdes(X), Ydes(Y))
                    Case 1
                        Dessin.PSet (Xdes(X), Ydes(Y))
                End Select
            End If
            Fenetre_Val.ForeColor = Fenetre_Val.BackColor
            Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
            Fenetre_Val.Print Space$(50)
            Fenetre_Val.Print Space$(50)
            Fenetre_Val.ForeColor = CouleurCou&
            Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
            Fenetre_Val.Print " X ="; X
            Fenetre_Val.Print " Y ="; Y
        End If
        SD% = SE%
        SE% = 0
    Loop
    Fenetre_Val.Cls
    DoFlag = False
    mnuTraceStop.Caption = "&Tracer"
    mnuAide.Enabled = True
    mnuDefinir.Enabled = True
    mnuEnregistrer.Enabled = True
    mnuImprimer.Enabled = True
    mnuOuvrir.Enabled = True
    Dessin.ForeColor = COULEUR_AVANT
    Fenetre_Prin.SetFocus
    Exit Sub
Traite_Erreurs0:
    Select Case Err
        Case 5, 6, 11
            ' 5  : invalid procedure call
            ' 6  : overflow
            ' 11 : division by zero
            SE% = 1
            Resume Next
        Case Else
            Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
            MsgBox Message$, 48, "COURBE"
            Exit Sub
    End Select
End Sub

Private Sub TraceCouCarExpl()
    '-----------------------------------------------------------------
    '-----------         Initialisation                ---------------
    '-----------------------------------------------------------------
    mnuAide.Enabled = False
    mnuDefinir.Enabled = False
    mnuEnregistrer.Enabled = False
    mnuImprimer.Enabled = False
    mnuOuvrir.Enabled = False
    '-----------------------------------------------------------------
    '-----------     Paramètres pour l'affichage       ---------------
    '-----------   non réglables par l'utilisateur     ---------------
    '-----------------------------------------------------------------
    LX% = Dessin.ScaleWidth
    LY% = Dessin.ScaleHeight
    ' Facteur d'échelle
    FE = LY% / LX%
    If Ech% = 1 Then
        Ymax = FE * (Xmax - Xmin) + Ymin
    End If
    '-----------------------------------------------------------------
    '----------- Déterminations de différentes valeurs ---------------
    '-----------             pour le tracé             ---------------
    '-----------------------------------------------------------------
    AX = Xmax - Xmin
    If AX < 0 Then
        MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
        Exit Sub
    End If
    If Ech% = 1 Then
        AY = AX * FE
        Ymax = Ymin + AY
    Else
        AY = Ymax - Ymin
        If AY < 0 Then
            MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
            Exit Sub
        End If
    End If
    MX = LX% / AX
    MY = -LY% / AY
    RX = 1 / MX
    RY = -1 / MY
    '---------------------------------------------------------------
    CouleurCou& = FenetreDefCourbe.pctColorCourbe.BackColor
    '---------------------------------------------------------------
    '-------------------------- Variables --------------------------
    '---------------------------------------------------------------
    NombVar% = 1
    var$(1) = "X"
    '--------------------------------------------------------------
    '--------- Première manipulation des chaines DefVar : ---------
    '--------- Suppression des blancs dans les formules,  ---------
    '--------- vérification du nombre de parenthèses et   ---------
    '--------- remplacement des constantes par leur valeur --------
    '--------------------------------------------------------------
    TextVar2$ = UCase$(TextVar2$)
    Call OteBlancs(TextVar2$, CorVar2$)
    Call Constante(CorVar2$, ForVar2$)
    If Erreur = True Then Exit Sub
    '--------------------------------------------------------------
    '--------- Effacer le dessin précédent ------------------------
    '--------------------------------------------------------------
    If Gap% = 1 Then
        Dessin.Picture = LoadPicture()
    End If
    On Error GoTo TraiteErreursCE
    '--------------------------------------------------------------
    '--------- Tracé du cadre et des axes -------------------------
    '--------------------------------------------------------------
    If Trac% = 1 Then
        Dessin.Line (0, 0)-(0, LY%)
        Dessin.Line -(LX%, LY%)
        Dessin.Line -(LX%, 0)
        Dessin.Line -(0, 0)
    End If
    If Tric% = 1 Then
        Dessin.Line (0, 0)-(0, LY%)
        Dessin.Line -(LX%, LY%)
        Dessin.Line -(LX%, 0)
        Dessin.Line -(0, 0)
    End If
    If Trax% = 1 Then
        Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
        Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
    End If
    '--------------------------------------------------------------
    '--------- Graduation et grille -------------------------------
    '--------------------------------------------------------------
    If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
    '--------------------------------------------------------------
    '--------- Tracé de la courbe ---------------------------------
    '--------------------------------------------------------------
    Dessin.ForeColor = CouleurCou&
    X = Xmin - RX
    SD% = 1
    SE% = 0
    Fenetre_Val.ForeColor = NOIR
    Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
    Fenetre_Val.Print " Un instant SVP ..."
    Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      X = X + RX
      If X > Xmax Then Exit Do
      vvar$(1) = LTrim$(Format(X, "0.000000"))
      Call variable(ForVar2$, ValVar2$)
      Call Traitement(ValVar2$, SorVar2$)
      Y = CSng(SorVar2$)
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If Y > Ymax Or Y < Ymin Then
            SE% = 2
            Select Case SD%
            Case 0
               Dessin.Line -(Xdes(X), Ydes(Y))
            Case 1, 2
               Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
         Else
            Select Case SD%
            Case 0, 2
               Dessin.Line -(Xdes(X), Ydes(Y))
            Case 1
               Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
    Loop
    Fenetre_Val.Cls
    DoFlag = False
    mnuTraceStop.Caption = "&Tracer"
    mnuAide.Enabled = True
    mnuDefinir.Enabled = True
    mnuEnregistrer.Enabled = True
    mnuImprimer.Enabled = True
    mnuOuvrir.Enabled = True
    Dessin.ForeColor = COULEUR_AVANT
    Fenetre_Prin.SetFocus
    Exit Sub
TraiteErreursCE:
    Select Case Err
        Case 5, 6, 11
            ' 5  : invalid procedure call
            ' 6  : overflow
            ' 11 : division by zero
            SE% = 1
            Resume Next
        Case Else
            Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
            MsgBox Message$, 48, "COURBE"
            Exit Sub
    End Select
End Sub

Private Sub TraceCouEqDif()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
      Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
      Exit Sub
   End If
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
      If AY < 0 Then
         MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
         Exit Sub
      End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '---------------------------------------------------------------
   '-------------------------- Variables --------------------------
   '---------------------------------------------------------------
   NombVar% = 2
   var$(1) = "X"
   var$(2) = "Y"
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   On Error GoTo Traite_ErreursED
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé éventuel des points --------------------------
   '--------------------------------------------------------------
   If Trap% = 1 And NbPts% > 0 Then
      LG% = Int((LX% + LY%) / 2)
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts%
         Dessin.Line (Xdes(Xpt(i%)) - 0.01 * LG%, Ydes(Ypt(i%)))-(Xdes(Xpt(i%)) + 0.01 * LG%, Ydes(Ypt(i%)))
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)) - 0.01 * LG%)-(Xdes(Xpt(i%)), Ydes(Ypt(i%)) + 0.01 * LG%)
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Reliure éventuelle des points --------------------------
   '--------------------------------------------------------------
   If Rep% = 1 And NbPts% > 0 Then
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts% - 1
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)))-(Xdes(Xpt(i% + 1)), Ydes(Ypt(i% + 1)))
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   '--------------------------------------------------------------
   ' 1) : tracé de (Xvnd, Yvnd) à Xmin
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xvnd
   Y = Yvnd
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((Xvnd - X) / AX * 100); "%    effectué"
      If X < Xmin Then Exit Do
      Call RungeKuttaMoins
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   '--------------------------------------------------------------
   ' 2) : tracé de (Xvnd, Yvnd) à Xmax
   '--------------------------------------------------------------
   X = Xvnd
   Y = Yvnd
   SD% = 1
   SE% = 0
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      If X > Xmax Then Exit Do
      Call RungeKuttaPlus
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_ErreursED:
   Select Case Err
      Case 5, 6, 11
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zero
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
      End Select
End Sub

Private Sub TraceCouEqDif2()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
      Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
      Exit Sub
   End If
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
      If AY < 0 Then
         MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
         Exit Sub
      End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '---------------------------------------------------------------
   '-------------------------- Variables --------------------------
   '---------------------------------------------------------------
   NombVar% = 3
   var$(1) = "X"
   var$(2) = "Y"
   ' dY/dX = Z
   var$(3) = "Z"
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   On Error GoTo Traite_ErreursED2
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé éventuel des points --------------------------
   '--------------------------------------------------------------
   If Trap% = 1 And NbPts% > 0 Then
      LG% = Int((LX% + LY%) / 2)
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts%
         Dessin.Line (Xdes(Xpt(i%)) - 0.01 * LG%, Ydes(Ypt(i%)))-(Xdes(Xpt(i%)) + 0.01 * LG%, Ydes(Ypt(i%)))
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)) - 0.01 * LG%)-(Xdes(Xpt(i%)), Ydes(Ypt(i%)) + 0.01 * LG%)
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Reliure éventuelle des points --------------------------
   '--------------------------------------------------------------
   If Rep% = 1 And NbPts% > 0 Then
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts% - 1
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)))-(Xdes(Xpt(i% + 1)), Ydes(Ypt(i% + 1)))
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   '--------------------------------------------------------------
   ' 1) : tracé de (Xvnd, Yvnd) à Xmin
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xvnd
   Y = Yvnd
   Z = Zvnd
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((Xvnd - X) / AX * 100); "%    effectué"
      If X < Xmin Then Exit Do
      Call EulerED2Moins
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   '--------------------------------------------------------------
   ' 2) : tracé de (Xvnd, Yvnd) à Xmax
   '--------------------------------------------------------------
   X = Xvnd
   Y = Yvnd
   Z = Zvnd
   SD% = 1
   SE% = 0
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      If X > Xmax Then Exit Do
      Call EulerED2Plus
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_ErreursED2:
   Select Case Err
      Case 5, 6, 11
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zero
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
      End Select
End Sub

Private Sub TraceCouInPoLag()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
       Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   If NbPtsI% < 2 Then
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "InPoLag"
   End If
   AX = Xmax - Xmin
   If AX < 0 Then
       MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
       Exit Sub
   End If
   If Ech% = 1 Then
       AY = AX * FE
       Ymax = Ymin + AY
   Else
       AY = Ymax - Ymin
       If AY < 0 Then
           MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
           Exit Sub
       End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
       Dessin.Picture = LoadPicture()
   End If
   On Error GoTo TraiteErreursPL
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then               ' tracer le cadre
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé des points --------------------------
   '--------------------------------------------------------------
   LG% = Int((LX% + LY%) / 2)
   Dessin.ForeColor = CouleurPts&
   For i% = 1 To NbPtsI%
      Dessin.Line (Xdes(Xpl(i%)) - 0.01 * LG%, Ydes(Ypl(i%)))-(Xdes(Xpl(i%)) + 0.01 * LG%, Ydes(Ypl(i%)))
      Dessin.Line (Xdes(Xpl(i%)), Ydes(Ypl(i%)) - 0.01 * LG%)-(Xdes(Xpl(i%)), Ydes(Ypl(i%)) + 0.01 * LG%)
   Next i%
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xmin - RX
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      X = X + RX
      If X > Xmax Then Exit Do
      ' ************************************************
      ' ********  routine de calcul de Y=Pn(X)  ********
      ' ************************************************
      Y = 0
      For i% = 0 To DegPoLag%
         Y = Y + UparI(i%) * X ^ i%
      Next i%
      ' ************************************************
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
TraiteErreursPL:
    Select Case Err
      Case 5, 6, 11
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zeroSE% = 1
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
      End Select
End Sub

Private Sub TraceCouInteg()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
      Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
      Exit Sub
   End If
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
      If AY < 0 Then
         MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
         Exit Sub
      End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '---------------------------------------------------------------
   '-------------------------- Variables --------------------------
   '---------------------------------------------------------------
   NombVar% = 1
   var$(1) = "X"
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   On Error GoTo TraiteErreursIN
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé éventuel des points --------------------------
   '--------------------------------------------------------------
   If Trap% = 1 And NbPts% > 0 Then
      LG% = Int((LX% + LY%) / 2)
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts%
         Dessin.Line (Xdes(Xpt(i%)) - 0.01 * LG%, Ydes(Ypt(i%)))-(Xdes(Xpt(i%)) + 0.01 * LG%, Ydes(Ypt(i%)))
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)) - 0.01 * LG%)-(Xdes(Xpt(i%)), Ydes(Ypt(i%)) + 0.01 * LG%)
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Reliure éventuelle des points --------------------------
   '--------------------------------------------------------------
   If Rep% = 1 And NbPts% > 0 Then
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts% - 1
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)))-(Xdes(Xpt(i% + 1)), Ydes(Ypt(i% + 1)))
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xmin - RX
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
   DoEvents
   If DoFlag = False Then Exit Do
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
   X = X + RX
   If X > Xmax Then Exit Do
   vvar$(1) = LTrim$(Format(X, "0.000000"))
   Call variable(ForVar2$, ValVar2$)
   Call Traitement(ValVar2$, SorVar2$)
   Y = CSng(SorVar2$)
   If Erreur = True Then Exit Do
   'SE%=0 : point dans le cadre
   'SE%=1 : point impossible à calculer, erreur
   'SE%=2 : point calculé mais hors cadre
   If SE% <> 1 Then
      If Y > Ymax Or Y < Ymin Then
         SE% = 2
         Select Case SD%
         Case 0
            Dessin.Line -(Xdes(X), Ydes(Y))
         Case 1, 2
            Dessin.PSet (Xdes(X), Ydes(Y))
         End Select
      Else
         Select Case SD%
         Case 0, 2
            Dessin.Line -(Xdes(X), Ydes(Y))
         Case 1
            Dessin.PSet (Xdes(X), Ydes(Y))
         End Select
      End If
      '*********************************************
      ' hachures verticales pour zone intégrée
      '*********************************************
      If X >= Xa And X <= Xb Then
         Dessin.Line (Xdes(X), Ydes(0))-(Xdes(X), Ydes(Y)), CouleurHach&
      End If
      '*********************************************
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = CouleurCou&
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print " X ="; X
      Fenetre_Val.Print " Y ="; Y
   End If
   SD% = SE%
   SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
TraiteErreursIN:
   Select Case Err
   Case 5, 6, 11
      ' 5  : invalid procedure call
      ' 6  : overflow
      ' 11 : division by zeroSE% = 1
      SE% = 1
      Resume Next
   Case Else
      Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
      MsgBox Message$, 48, "COURBE"
      Exit Sub
   End Select
End Sub

Private Sub TraceCouMoinCar()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
       Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   If NbPtsMoinCar% < NbPar% Then
      MsgBox "Le nombre de points doit être supérieur ou égal au nombre de paramètres !", 48, "COURBE"
   End If
   AX = Xmax - Xmin
   If AX < 0 Then
       MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
       Exit Sub
   End If
   If Ech% = 1 Then
       AY = AX * FE
       Ymax = Ymin + AY
   Else
       AY = Ymax - Ymin
       If AY < 0 Then
           MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
           Exit Sub
       End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '-------------------------- Variables --------------------------
   NombVar% = 1
   var$(1) = "X"
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   Call Parametre(ForVar2$, ParVar2$)
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
       Dessin.Picture = LoadPicture()
   End If
   On Error GoTo TraiteErreursMC
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then               ' tracer le cadre
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé des points --------------------------
   '--------------------------------------------------------------
   LG% = Int((LX% + LY%) / 2)
   Dessin.ForeColor = CouleurPts&
   For i% = 1 To NbPtsMoinCar%
      Dessin.Line (Xdes(Xmc(i%)) - 0.01 * LG%, Ydes(Ymc(i%)))-(Xdes(Xmc(i%)) + 0.01 * LG%, Ydes(Ymc(i%)))
      Dessin.Line (Xdes(Xmc(i%)), Ydes(Ymc(i%)) - 0.01 * LG%)-(Xdes(Xmc(i%)), Ydes(Ymc(i%)) + 0.01 * LG%)
   Next i%
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xmin - RX
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(60)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      X = X + RX
      If X > Xmax Then Exit Do
      vvar$(1) = LTrim$(Format(X, "0.000000"))
      Call variable(ParVar2$, ValVar2$)
      Call Traitement(ValVar2$, SorVar2$)
      Y = CSng(SorVar2$)
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuDefinir.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
TraiteErreursMC:
   Select Case Err
      Case 5, 6, 11
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zero
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
   End Select
End Sub

Private Sub TraceCouPolynome()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   Fenetre_Val.Cls
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   Ymax = FE * (Xmax - Xmin) + Ymin
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
      Exit Sub
   End If
   AY = AX * FE
   Ymax = Ymin + AY
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   On Error GoTo Traite_ErreursPO
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé éventuel des points --------------------------
   '--------------------------------------------------------------
   If Trap% = 1 And NbPts% > 0 Then
      LG% = Int((LX% + LY%) / 2)
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts%
         Dessin.Line (Xdes(Xpt(i%)) - 0.01 * LG%, Ydes(Ypt(i%)))-(Xdes(Xpt(i%)) + 0.01 * LG%, Ydes(Ypt(i%)))
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)) - 0.01 * LG%)-(Xdes(Xpt(i%)), Ydes(Ypt(i%)) + 0.01 * LG%)
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Reliure éventuelle des points --------------------------
   '--------------------------------------------------------------
   If Rep% = 1 And NbPts% > 0 Then
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts% - 1
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)))-(Xdes(Xpt(i% + 1)), Ydes(Ypt(i% + 1)))
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xmin - RX
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      X = X + RX
      If X > Xmax Then Exit Do
      ' *************************************************
      '                   calcul de Y
      ' *************************************************
      Y = 0
      For i% = 0 To DegPpol%
         Y = Y + Ppol(i%) * X ^ i%
      Next i%
      ' *************************************************
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_ErreursPO:
   Select Case Err
   Case 5, 6, 11
      ' 5  : invalid procedure call
      ' 6  : overflow
      ' 11 : division by zero
      SE% = 1
      Resume Next
   Case Else
      Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
      MsgBox Message$, 48, "COURBE"
      Exit Sub
   End Select
End Sub

Private Sub TraceCouPrimitive()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
      Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
      Exit Sub
   End If
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
      If AY < 0 Then
         MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
         Exit Sub
      End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '---------------------------------------------------------------
   '-------------------------- Variables --------------------------
   '---------------------------------------------------------------
   NombVar% = 1
   var$(1) = "X"
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   On Error GoTo TraiteErreursPrim
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé éventuel des points --------------------------
   '--------------------------------------------------------------
   If Trap% = 1 And NbPts% > 0 Then
      LG% = Int((LX% + LY%) / 2)
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts%
         Dessin.Line (Xdes(Xpt(i%)) - 0.01 * LG%, Ydes(Ypt(i%)))-(Xdes(Xpt(i%)) + 0.01 * LG%, Ydes(Ypt(i%)))
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)) - 0.01 * LG%)-(Xdes(Xpt(i%)), Ydes(Ypt(i%)) + 0.01 * LG%)
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Reliure éventuelle des points --------------------------
   '--------------------------------------------------------------
   If Rep% = 1 And NbPts% > 0 Then
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts% - 1
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)))-(Xdes(Xpt(i% + 1)), Ydes(Ypt(i% + 1)))
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   '--------------------------------------------------------------
   ' 1) : tracé de Xvnd à Xmin
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xvnd
   Y = 0
   YA = 0
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
   DoEvents
   If DoFlag = False Then Exit Do
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Tracé :  "; Int((Xvnd - X) / AX * 100); "%    effectué"
   If X < Xmin Then Exit Do
   Call SimpsonPrimMoins
   If Erreur = True Then Exit Do
   'SE%=0 : point dans le cadre
   'SE%=1 : point impossible à calculer, erreur
   'SE%=2 : point calculé mais hors cadre
   If SE% <> 1 Then
      If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
            SE% = 2
            Select Case SD%
               Case 0
                  Dessin.Line -(Xdes(X), Ydes(Y))
               Case 1, 2
                  Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
      Else
            Select Case SD%
               Case 0, 2
                  Dessin.Line -(Xdes(X), Ydes(Y))
               Case 1
                  Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
      End If
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = CouleurCou&
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print " X ="; X
      Fenetre_Val.Print " Y ="; Y
   End If
   SD% = SE%
   SE% = 0
   Loop
   '--------------------------------------------------------------
   ' 2) : tracé de Xvnd à Xmax
   '--------------------------------------------------------------
   X = Xvnd
   Y = 0
   YA = 0
   SD% = 1
   SE% = 0
   Do
   DoEvents
   If DoFlag = False Then Exit Do
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
   If X > Xmax Then Exit Do
   Call SimpsonPrimPlus
   If Erreur = True Then Exit Do
   'SE%=0 : point dans le cadre
   'SE%=1 : point impossible à calculer, erreur
   'SE%=2 : point calculé mais hors cadre
   If SE% <> 1 Then
      If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
            SE% = 2
            Select Case SD%
               Case 0
                  Dessin.Line -(Xdes(X), Ydes(Y))
               Case 1, 2
                  Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
      Else
            Select Case SD%
               Case 0, 2
                  Dessin.Line -(Xdes(X), Ydes(Y))
               Case 1
                  Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
      End If
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = CouleurCou&
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print " X ="; X
      Fenetre_Val.Print " Y ="; Y
   End If
   SD% = SE%
   SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
TraiteErreursPrim:
   Select Case Err
      Case 5, 6, 11
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zero
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
   End Select
End Sub
   

Private Sub TraceCourbe()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
      Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   AX = Xmax - Xmin
   If Ech% = 1 Then
      AY = AX * FE
      Ymax = Ymin + AY
   Else
      AY = Ymax - Ymin
   End If
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin", 48, "COURBE"
      Exit Sub
   ElseIf AY < 0 Then
      MsgBox "Ymax doit être supérieur à Ymin", 48, "COURBE"
      Exit Sub
   ElseIf Tmax < Tmin Then
      MsgBox "Tmax doit être supérieur à Tmin", 48, "COURBE"
      Exit Sub
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '---------------------------------------------------------------
   '-------------------------- Variables --------------------------
   '---------------------------------------------------------------
   Select Case Expr%
      Case 1
         NombVar% = 1
         Select Case Coor%
               Case 1
                  var$(1) = "X"
               Case 2
                  var$(1) = "TETA"
         End Select
      Case 2
         NombVar% = 2
         Select Case Coor%
               Case 1
                  var$(1) = "X"
                  var$(2) = "Y"
               Case 2
                  var$(1) = "TETA"
                  var$(2) = "R"
         End Select
      Case 3
         NombVar% = 1
         var$(1) = "T"
   End Select
   '--------------------------------------------------------------
   '--------- Première manipulation des chaines DefVar : ---------
   '--------- Suppression des blancs dans les formules,  ---------
   '--------- vérification du nombre de parenthèses et   ---------
   '--------- remplacement des constantes par leur valeur --------
   '--------------------------------------------------------------
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
   If Expr% = 3 Then
      Call OteBlancs(TextVar1$, CorVar1$)
      Call Constante(CorVar1$, ForVar1$)
   End If
   If Erreur = True Then Exit Sub
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   '--------------------------------------------------------------
   '--------- Traitement des erreurs -----------------------------
   '--------------------------------------------------------------
   On Error GoTo Traite_Erreurs
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé éventuel des points --------------------------
   '--------------------------------------------------------------
   If Trap% = 1 And NbPts% > 0 Then
      LG% = Int((LX% + LY%) / 2)
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts%
         Dessin.Line (Xdes(Xpt(i%)) - 0.01 * LG%, Ydes(Ypt(i%)))-(Xdes(Xpt(i%)) + 0.01 * LG%, Ydes(Ypt(i%)))
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)) - 0.01 * LG%)-(Xdes(Xpt(i%)), Ydes(Ypt(i%)) + 0.01 * LG%)
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Reliure éventuelle des points --------------------------
   '--------------------------------------------------------------
   If Rep% = 1 And NbPts% > 0 Then
      Dessin.ForeColor = CouleurPts&
      For i% = 1 To NbPts% - 1
         Dessin.Line (Xdes(Xpt(i%)), Ydes(Ypt(i%)))-(Xdes(Xpt(i% + 1)), Ydes(Ypt(i% + 1)))
      Next i%
   End If
   '--------------------------------------------------------------
   '--------- Tracé éventuel de la tangente ----------------------
   '--------------------------------------------------------------
   If Tran% = 1 Then
      Dessin.ForeColor = CouleurTan&
      '-------------------------------------------------------
      ' calcul de la tangente (sauf pour expression implicite)
      '-------------------------------------------------------
      tanvert = False
      epsloc = 0.00001
      Select Case Expr%
      Case 1 'explicite
        ' -----------------------------------------
        ' calcul de la dérivée F'(Xo) ou R'(TETAo)
        ' -----------------------------------------
        X = Xpar
        vvar$(1) = LTrim$(Format(X, "0.000000"))
        Call variable(ForVar2$, ValVar2$)
        Call Traitement(ValVar2$, SorVar2$)
        Ypar = CSng(SorVar2$)
        '---
        X1 = Xpar + epsloc
        X = X1
        vvar$(1) = LTrim$(Format(X, "0.000000"))
        Call variable(ForVar2$, ValVar2$)
        Call Traitement(ValVar2$, SorVar2$)
        Y1 = CSng(SorVar2$)
        '---
        fprim = (Y1 - Ypar) / epsloc    '  F'(Xo)
        ' --------------------------------
        ' tangente : y = atan * x + btan :
        ' --------------------------------
        Select Case Coor%
            Case 1
                ' ------------------------
                ' coordonnées cartésiennes
                ' ------------------------
                ' Y = F'(Xo)*X - F'(Xo)*Xo + Yo
                atan = fprim
                btan = Ypar - atan * Xpar
            Case 2
                ' --------------------
                ' coordonnées polaires
                ' --------------------
                ' denom = R'(TETAo)*cos(TETAo) - R(TETAo)*sin(TETAo)
                ' Y = (1/denom)*{[R'(TETAo)*sin(TETAo)+R(TETAo)*cos(TETAo)]*X-R²(TETAo)}
                denom = fprim * Cos(Xpar) - Ypar * Sin(Xpar)
                If Abs(denom) < epsloc Then 'tangente verticale
                    tanvert = True
                Else
                    atan = (fprim * Sin(Xpar) + Ypar * Cos(Xpar)) / denom
                    btan = -Ypar * Ypar / denom
                End If
        End Select
        ' -----------------------------------------
        ' recherche des points d'intersection de la
        ' tangente avec le cadre pour la tracer :
        ' -----------------------------------------
        If tanvert = True Then  'tangente verticale
            Xmit = Ypar * Cos(Xpar)
            Ymit = Ymin
            Xmat = Xmit
            Ymat = Ymax
        Else
            Ymit = atan * Xmin + btan
            If Ymit < Ymin Then
                Ymit = Ymin
                Xmit = (Ymit - btan) / atan
            ElseIf Ymit > Ymax Then
                Ymit = Ymax
                Xmit = (Ymit - btan) / atan
            Else
                Xmit = Xmin
            End If
            Ymat = atan * Xmax + btan
            If Ymat > Ymax Then
                Ymat = Ymax
                Xmat = (Ymat - btan) / atan
            ElseIf Ymat < Ymin Then
                Ymat = Ymin
                Xmat = (Ymat - btan) / atan
            Else
                Xmat = Xmax
            End If
        End If
        Dessin.Line (Xdes(Xmit), Ydes(Ymit))-(Xdes(Xmat), Ydes(Ymat))
      Case 2  ' implicite
        ' ---------------------------------------------
        ' calcul effectué plus loin, au moment du tracé
        ' ---------------------------------------------
      Case 3  ' paramétrée
        ' -----------------------------------------
        ' calcul des dérivées dX/dT(To) et dY/dT(To)
        ' -----------------------------------------
        T = Xpar
        vvar$(1) = LTrim$(Format(T, "0.000000"))
        Call variable(ForVar1$, ValVar1$)
        Call variable(ForVar2$, ValVar2$)
        Call Traitement(ValVar1$, SorVar1$)
        Call Traitement(ValVar2$, SorVar2$)
        X = CSng(SorVar1$)
        Y = CSng(SorVar2$)
        '---
        T = Xpar + epsloc
        vvar$(1) = LTrim$(Format(T, "0.000000"))
        Call variable(ForVar1$, ValVar1$)
        Call variable(ForVar2$, ValVar2$)
        Call Traitement(ValVar1$, SorVar1$)
        Call Traitement(ValVar2$, SorVar2$)
        X1 = CSng(SorVar1$)
        Y1 = CSng(SorVar2$)
        '---
        fprimX = (X1 - X) / epsloc  ' dX/dT(To)
        fprimY = (Y1 - Y) / epsloc  ' dY/dT(To)
        ' ----------------------------------------------------------------------
        ' tangente : Y = [dY/dT(To)]/[dX/dT(To)]*X-[dY/dT(To)]/[dX/dT(To)]*Xo+Yo
        ' tangente : y = atan * x + btan :
        ' ----------------------------------------------------------------------
        Select Case Coor%
            Case 1 ' coordonnées cartésiennes
                If Abs(fprimX) < epsloc Then 'tangente verticale
                    tanvert = True
                Else
                    atan = fprimY / fprimX
                    btan = Y - atan * X
                End If
            Case 2 ' coordonnées polaires
                Message$ = "calcul de la tangente non prévu dans ce cas"
                MsgBox Message$, 48, "COURBE"
        End Select
        ' -----------------------------------------
        ' recherche des points d'intersection de la
        ' tangente avec le cadre pour la tracer :
        ' -----------------------------------------
        If tanvert = True Then  'tangente verticale
            Xmit = X
            Ymit = Ymin
            Xmat = Xmit
            Ymat = Ymax
        Else
            Ymit = atan * Xmin + btan
            If Ymit < Ymin Then
                Ymit = Ymin
                Xmit = (Ymit - btan) / atan
            ElseIf Ymit > Ymax Then
                Ymit = Ymax
                Xmit = (Ymit - btan) / atan
            Else
                Xmit = Xmin
            End If
            Ymat = atan * Xmax + btan
            If Ymat > Ymax Then
                Ymat = Ymax
                Xmat = (Ymat - btan) / atan
            ElseIf Ymat < Ymin Then
                Ymat = Ymin
                Xmat = (Ymat - btan) / atan
            Else
                Xmat = Xmax
            End If
        End If
        If tanvert = True And Abs(fprimY) < epsloc Then
            Message$ = "tangente en un point stationnaire"
            MsgBox Message$, 48, "COURBE"
        Else
            Dessin.Line (Xdes(Xmit), Ydes(Ymit))-(Xdes(Xmat), Ydes(Ymat))
        End If
      End Select
   End If
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   Select Case Expr%
      Case 1
         Select Case Coor%
               Case 1
                  Imax% = LX%
               Case 2
                  Imax% = 800
         End Select
      Case 2
         Imax% = LX%
         Jmax% = LY%
      Case 3
         Imax% = 800
   End Select
   i% = -1
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
   DoEvents
   If DoFlag = False Then Exit Do
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Tracé :  "; Int(i% / Imax% * 100); "%    effectué"
   i% = i% + 1
   If i% > Imax% Then Exit Do
   Select Case Expr%
      Case 1
            Select Case Coor%
               Case 1
                  X = Xmin + RX * i%
                  vvar$(1) = LTrim$(Format(X, "0.000000"))
                  Call variable(ForVar2$, ValVar2$)
                  Call Traitement(ValVar2$, SorVar2$)
                  Y = CSng(SorVar2$)
               Case 2
                  TETA = 2 * PI * i% / Imax%
                  vvar$(1) = LTrim$(Format(TETA, "0.000000"))
                  Call variable(ForVar2$, ValVar2$)
                  Call Traitement(ValVar2$, SorVar2$)
                  R = CSng(SorVar2$)
                  X = R * Cos(TETA)
                  Y = R * Sin(TETA)
            End Select
      Case 2
            X = Xmin + RX * i%
            S1 = 2
            j% = 0
            Do
               j% = j% + 1
               If j% > Jmax% Then Exit Do
               Y = Ymin + RY * j%
               Select Case Coor%
                  Case 1
                        vvar$(1) = LTrim$(Format(X, "0.000000"))
                        vvar$(2) = LTrim$(Format(Y, "0.000000"))
                  Case 2
                        TETA = Atn(Y / X)
                        If X < 0 Then
                           TETA = TETA + PI
                        ElseIf Y < 0 Then
                           TETA = 2 * PI + TETA
                        End If
                        R = Sqr(X * X + Y * Y)
                        vvar$(1) = LTrim$(Format(X, "0.000000"))
                        vvar$(2) = LTrim$(Format(Y, "0.000000"))
               End Select
               Call variable(ForVar2$, ValVar2$)
               Call Traitement(ValVar2$, SorVar2$)
               Z = CSng(SorVar2$)
               S2 = Sgn(Z)
               If S2 = -S1 Then
                  Dessin.PSet (Xdes(X), Ydes(Y))
                  If Abs(X - Xpar) < 0.6 * RX Then
                    If Tran% = 1 Then
                        '-----------------------------
                        ' calcul de la tangente
                        '-----------------------------
                        Select Case Coor%
                            Case 1 ' coordonnées cartésiennes
                                ' ----------------------------------
                                ' 1) calcul de la dérivée dF/dX (Xo)
                                ' ----------------------------------
                                X1 = X + epsloc
                                vvar$(1) = LTrim$(Format(X1, "0.000000"))
                                vvar$(2) = LTrim$(Format(Y, "0.000000"))
                                Call variable(ForVar2$, ValVar2$)
                                Call Traitement(ValVar2$, SorVar2$)
                                ZX = CSng(SorVar2$)
                                fprimX = (ZX - Z) / epsloc  ' dF/dX(Xo)
                                ' ----------------------------------
                                ' 2) calcul de la dérivée dF/dY (Xo)
                                ' ----------------------------------
                                Y1 = Y + epsloc
                                vvar$(1) = LTrim$(Format(X, "0.000000"))
                                vvar$(2) = LTrim$(Format(Y1, "0.000000"))
                                Call variable(ForVar2$, ValVar2$)
                                Call Traitement(ValVar2$, SorVar2$)
                                ZY = CSng(SorVar2$)
                                fprimY = (ZY - Z) / epsloc  ' dF/dY(Yo)
                                ' ----------------------------------------------------
                                ' tangente : [dF/dX (Xo)](X-Xo)+ [dF/dY (Yo)](X-Xo)= 0
                                ' tangente : y = atan * x + btan :
                                ' ----------------------------------------------------
                                If Abs(fprimY) < epsloc Then 'tangente verticale
                                    tanvert = True
                                Else
                                    atan = -fprimX / fprimY
                                    btan = Y - atan * X
                                End If
                                ' -----------------------------------------
                                ' recherche des points d'intersection de la
                                ' tangente avec le cadre pour la tracer :
                                ' -----------------------------------------
                                If tanvert = True Then  'tangente verticale
                                    Xmit = Xpar
                                    Ymit = Ymin
                                    Xmat = Xmit
                                    Ymat = Ymax
                                Else
                                    Ymit = atan * Xmin + btan
                                    If Ymit < Ymin Then
                                        Ymit = Ymin
                                        Xmit = (Ymit - btan) / atan
                                    ElseIf Ymit > Ymax Then
                                        Ymit = Ymax
                                        Xmit = (Ymit - btan) / atan
                                    Else
                                        Xmit = Xmin
                                    End If
                                    Ymat = atan * Xmax + btan
                                    If Ymat > Ymax Then
                                        Ymat = Ymax
                                        Xmat = (Ymat - btan) / atan
                                    ElseIf Ymat < Ymin Then
                                        Ymat = Ymin
                                        Xmat = (Ymat - btan) / atan
                                    Else
                                        Xmat = Xmax
                                    End If
                                End If
                                Dessin.Line (Xdes(Xmit), Ydes(Ymit))-(Xdes(Xmat), Ydes(Ymat))
                            Case 2 ' coordonnées polaires
                                Message$ = "calcul de la tangente non prévu dans ce cas"
                                MsgBox Message$, 48, "COURBE"
                        End Select
                    End If
                  End If
               End If
               S1 = S2
            Loop
            SE% = 1
      Case 3
            T = Tmin + (Tmax - Tmin) * i% / Imax%
            vvar$(1) = LTrim$(Format(T, "0.000000"))
            Call variable(ForVar1$, ValVar1$)
            Call variable(ForVar2$, ValVar2$)
            Call Traitement(ValVar1$, SorVar1$)
            Call Traitement(ValVar2$, SorVar2$)
            Select Case Coor%
               Case 1
                  X = CSng(SorVar1$)
                  Y = CSng(SorVar2$)
               Case 2
                  TETA = CSng(SorVar1$)
                  R = CSng(SorVar2$)
                  X = R * Cos(TETA)
                  Y = R * Sin(TETA)
            End Select
   End Select
   If Erreur = True Then Exit Do
   'SE%=0 : point dans le cadre
   'SE%=1 : point impossible à calculer, erreur
   'SE%=2 : point calculé mais hors cadre
   If SE% <> 1 Then
      If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
            SE% = 2
            Select Case SD%
               Case 0
                  Dessin.Line -(Xdes(X), Ydes(Y))
               Case 1, 2
                  Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
      Else
            Select Case SD%
               Case 0, 2
                  Dessin.Line -(Xdes(X), Ydes(Y))
               Case 1
                  Dessin.PSet (Xdes(X), Ydes(Y))
            End Select
      End If
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = CouleurCou&
      Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
      Fenetre_Val.Print " X ="; X
      Fenetre_Val.Print " Y ="; Y
   End If
   SD% = SE%
   SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_Erreurs:
   Select Case Err
      Case 5, 6, 11, 13
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zero
         ' 13 : type mismatch
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
   End Select
End Sub

Private Sub TraceCouRegreLin()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   Fenetre_Val.Cls
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   Ymax = FE * (Xmax - Xmin) + Ymin
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   If NbPtsReLin% < 2 Then
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "COURBE"
   End If
   AX = Xmax - Xmin
   If AX < 0 Then
      MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
      Exit Sub
   End If
   AY = AX * FE
   Ymax = Ymin + AY
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
      Dessin.Picture = LoadPicture()
   End If
   On Error GoTo Traite_ErreursRL
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   
   If Trac% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé des points --------------------------
   '--------------------------------------------------------------
   LG% = Int((LX% + LY%) / 2)
   Dessin.ForeColor = CouleurPts&
   For i% = 1 To NbPtsReLin%
      Dessin.Line (Xdes(Xrl(i%)) - 0.01 * LG%, Ydes(Yrl(i%)))-(Xdes(Xrl(i%)) + 0.01 * LG%, Ydes(Yrl(i%)))
      Dessin.Line (Xdes(Xrl(i%)), Ydes(Yrl(i%)) - 0.01 * LG%)-(Xdes(Xrl(i%)), Ydes(Yrl(i%)) + 0.01 * LG%)
   Next i%
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   X = Xmin - RX
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = NOIR
   Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
   Fenetre_Val.Print " Un instant SVP ..."
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.ForeColor = Fenetre_Val.BackColor
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - Xmin) / AX * 100); "%    effectué"
      X = X + RX
      If X > Xmax Then Exit Do
      ' *************************************************
      '                   calcul de Y
      ' *************************************************
      Y = Adroite * X + Bdroite
      ' *************************************************
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
Traite_ErreursRL:
   Select Case Err
   Case 5, 6, 11
      ' 5  : invalid procedure call
      ' 6  : overflow
      ' 11 : division by zero
      SE% = 1
      Resume Next
   Case Else
      Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
      MsgBox Message$, 48, "COURBE"
      Exit Sub
   End Select
End Sub

Private Sub TraceCouSpline()
   '-----------------------------------------------------------------
   '-----------         Initialisation                ---------------
   '-----------------------------------------------------------------
   mnuAide.Enabled = False
   mnuEnregistrer.Enabled = False
   mnuImprimer.Enabled = False
   mnuOuvrir.Enabled = False
   '-----------------------------------------------------------------
   '-----------     Paramètres pour l'affichage       ---------------
   '-----------   non réglables par l'utilisateur     ---------------
   '-----------------------------------------------------------------
   LX% = Dessin.ScaleWidth
   LY% = Dessin.ScaleHeight
   ' Facteur d'échelle
   FE = LY% / LX%
   If Ech% = 1 Then
       Ymax = FE * (Xmax - Xmin) + Ymin
   End If
   '-----------------------------------------------------------------
   '----------- Déterminations de différentes valeurs ---------------
   '-----------             pour le tracé             ---------------
   '-----------------------------------------------------------------
   If NbPtsSp% < 2 Then
      MsgBox "Le nombre de points doit être supérieur ou égal à 2 !", 48, "InPoLag"
   End If
   AX = Xmax - Xmin
   If AX < 0 Then
       MsgBox "Xmax doit être supérieur à Xmin !", 48, "COURBE"
       Exit Sub
   End If
   If Ech% = 1 Then
       AY = AX * FE
       Ymax = Ymin + AY
   Else
       AY = Ymax - Ymin
       If AY < 0 Then
           MsgBox "Ymax doit être supérieur à Ymin !", 48, "COURBE"
           Exit Sub
       End If
   End If
   MX = LX% / AX
   MY = -LY% / AY
   RX = 1 / MX
   RY = -1 / MY
   '--------------------------------------------------------------
   '--------- Effacer le dessin précédent ------------------------
   '--------------------------------------------------------------
   If Gap% = 1 Then
       Dessin.Picture = LoadPicture()
   End If
   On Error GoTo TraiteErreursSp
   '--------------------------------------------------------------
   '--------- Tracé du cadre et des axes -------------------------
   '--------------------------------------------------------------
   If Trac% = 1 Then               ' tracer le cadre
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Tric% = 1 Then
      Dessin.Line (0, 0)-(0, LY%)
      Dessin.Line -(LX%, LY%)
      Dessin.Line -(LX%, 0)
      Dessin.Line -(0, 0)
   End If
   If Trax% = 1 Then
      Dessin.Line (0, -Ymax * MY)-(LX%, -Ymax * MY)
      Dessin.Line (-Xmin * MX, 0)-(-Xmin * MX, LY%)
   End If
   '--------------------------------------------------------------
   '--------- Graduation et grille -------------------------------
   '--------------------------------------------------------------
   If Grax% = 1 Or Grac% = 1 Or Tric% = 1 Then Call GraduAx
   '--------------------------------------------------------------
   '--------- Tracé des points -----------------------------------
   '--------------------------------------------------------------
   LG% = Int((LX% + LY%) / 2)
   Dessin.ForeColor = CouleurPts&
   For i% = 1 To NbPtsSp%
      Dessin.Line (Xdes(Xsp(i%)) - 0.01 * LG%, Ydes(Ysp(i%)))-(Xdes(Xsp(i%)) + 0.01 * LG%, Ydes(Ysp(i%)))
      Dessin.Line (Xdes(Xsp(i%)), Ydes(Ysp(i%)) - 0.01 * LG%)-(Xdes(Xsp(i%)), Ydes(Ysp(i%)) + 0.01 * LG%)
   Next i%
   '--------------------------------------------------------------
   '--------- Tracé de la courbe ---------------------------------
   '--------------------------------------------------------------
   Dessin.ForeColor = CouleurCou&
   XminSp = Xsp(1)
   If Xmin > XminSp Then XminSp = Xmin
   XmaxSp = Xsp(NbPtsSp%)
   If Xmax < XmaxSp Then XmaxSp = Xmax
   AXSp = XmaxSp - XminSp
   X = XminSp - RX
   ISp% = 1
   SD% = 1
   SE% = 0
   Fenetre_Val.ForeColor = Fenetre_Val.BackColor
   Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
   Fenetre_Val.Print Space$(50)
   Do
      DoEvents
      If DoFlag = False Then Exit Do
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print Space$(50)
      Fenetre_Val.ForeColor = NOIR
      Fenetre_Val.PSet (0, 2500), Fenetre_Val.BackColor
      Fenetre_Val.Print " Tracé :  "; Int((X - XminSp) / AXSp * 100); "%    effectué"
      X = X + RX
      If X > XmaxSp Then Exit Do
      ' ************************************************
      ' ***********  routine de calcul de Y  ***********
      ' ************************************************
      If X > Xsp(ISp% + 1) Then
         ISp% = ISp% + 1
      End If
      DXsp = X - Xsp(ISp%)
      Y = Aspline(ISp%) * DXsp ^ 3 + Bspline(ISp%) * DXsp ^ 2 + Cspline(ISp%) * DXsp + Dspline(ISp%)
      ' ************************************************
      If Erreur = True Then Exit Do
      'SE%=0 : point dans le cadre
      'SE%=1 : point impossible à calculer, erreur
      'SE%=2 : point calculé mais hors cadre
      If SE% <> 1 Then
         If X > Xmax Or X < Xmin Or Y > Ymax Or Y < Ymin Then
               SE% = 2
               Select Case SD%
                  Case 0
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1, 2
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         Else
               Select Case SD%
                  Case 0, 2
                     Dessin.Line -(Xdes(X), Ydes(Y))
                  Case 1
                     Dessin.PSet (Xdes(X), Ydes(Y))
               End Select
         End If
         Fenetre_Val.ForeColor = Fenetre_Val.BackColor
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.Print Space$(50)
         Fenetre_Val.ForeColor = CouleurCou&
         Fenetre_Val.PSet (0, 3000), Fenetre_Val.BackColor
         Fenetre_Val.Print " X ="; X
         Fenetre_Val.Print " Y ="; Y
      End If
      SD% = SE%
      SE% = 0
   Loop
   Fenetre_Val.Cls
   DoFlag = False
   mnuTraceStop.Caption = "&Tracer"
   mnuAide.Enabled = True
   mnuEnregistrer.Enabled = True
   mnuImprimer.Enabled = True
   mnuOuvrir.Enabled = True
   Dessin.ForeColor = COULEUR_AVANT
   Fenetre_Prin.SetFocus
   Exit Sub
TraiteErreursSp:
   Select Case Err
      Case 5, 6, 11
         ' 5  : invalid procedure call
         ' 6  : overflow
         ' 11 : division by zero
         SE% = 1
         Resume Next
      Case Else
         Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
         MsgBox Message$, 48, "COURBE"
         Exit Sub
   End Select
End Sub

Private Sub ValGradX(X, Y)
    Rem ----- Routine valeurs des graduations -----
    Rem <<< graduations en x >>>
    acar$ = Format(X)
    lv = Len(acar$)
    Dessin.PSet (Xdes(X) - lv / 2, Ydes(Y)), Dessin.BackColor
    Dessin.Print acar$
End Sub

Private Sub ValGradY(X, Y)
    Rem ----- Routine valeurs des graduations -----
    Rem <<< graduations en y >>>
    Dessin.PSet (Xdes(X), Ydes(Y)), Dessin.BackColor
    Dessin.Print Format(Y)
End Sub

Private Function Xdes(X)
    Xdes = (X - Xmin) * MX
End Function

Private Function Ydes(Y)
    Ydes = (Y - Ymax) * MY
End Function


Public Sub TraceStop()
   Select Case DoFlag
   Case True
      mnuTraceStop.Caption = "&Tracer"
      DoFlag = False
   Case False
      mnuTraceStop.Caption = "&Arrêter le tracé"
      DoFlag = True
      Select Case NumChoisi%
      Case NUMINTEGRALE
         TraceCouInteg
      Case NUMCOUR, NUMSOLEQ
         TraceCourbe
      Case NUMEQUADIF
         TraceCouEqDif
      Case NUMEQUADIF2
         TraceCouEqDif2
      Case NUMPRIMITIVE
         TraceCouPrimitive
      Case NUMOINCAR
         TraceCouMoinCar
      Case NUMREGRELIN
         TraceCouRegreLin
      Case NUMINPOLAG
         TraceCouInPoLag
      Case NUMSPLINE
         TraceCouSpline
      Case NUMPOLYNOME
         TraceCouPolynome
      End Select
   End Select
End Sub

Private Sub mnuRetour_Click()
   Fenetre_Prin.Hide
   Fenetre_Val.Hide
   Dessin.Hide
End Sub

Public Sub mnuTraceStop_Click()
   Call TraceStop
End Sub


