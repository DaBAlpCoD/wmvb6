VERSION 5.00
Begin VB.Form FenetreVisuS3D 
   Caption         =   "Surface 3D : visualiser"
   ClientHeight    =   7410
   ClientLeft      =   330
   ClientTop       =   750
   ClientWidth     =   10080
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
   ScaleHeight     =   7410
   ScaleWidth      =   10080
   Begin VB.CheckBox chkGarder 
      Caption         =   "Garder tracé précédent"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.PictureBox pctVueS3D 
      AutoRedraw      =   -1  'True
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label lblFormule 
      Height          =   735
      Left            =   7440
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblMessage 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu mnuEnregistrer 
         Caption         =   "&Enregistrer..."
      End
   End
   Begin VB.Menu mnuTraceStop 
      Caption         =   "&Tracer"
   End
   Begin VB.Menu mnuRetour 
      Caption         =   "&Retour"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "FenetreVisuS3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MX%(), MI%()
Dim DES3D%     ' largeur et hauteur de pctVueS3D (en twips)
Dim MS3D
Dim Xs3d, Ys3d, Zs3d
Dim Xa, YA, XM, YM, XP, YP
Dim XE, YE, XX, YY
Dim CX, CY
Dim EcartXm, EcartYm, EcartXp, EcartYp
Dim MXs3d
Dim SigneZ%
Dim is3d%, js3d%, ks3d%
Dim NbLiS3D%, LongPasS3D%, NbPasS3D%
Dim Direction%
Dim Sym%
Dim LigneVisibleFlag%
Dim LanceTrace%
Dim FermeVisuS3D%

Private Sub CalculeS3D()
' ******************************************
' Surface d'une fonction Z=F(X,Y) en 3D
' Théoric n°25 p 32
' ******************************************
' *******************
' programme principal
' *******************
' Effacer le dessin précédent
If Gap3D% = 0 Then
   pctVueS3D.Picture = LoadPicture()
End If
' Nombre de lignes
NbLiS3D% = 20
' Nombre de pas entre les lignes
LongPasS3D% = 15
NbPasS3D% = NbLiS3D% * LongPasS3D%
ReDim MX%(2 * NbPasS3D%), MI%(2 * NbPasS3D%)
' *****************************************
Call PreTraiteFonction
If Erreur = True Then Exit Sub
' *****************************************
XE = DES3D%
YE = DES3D%
MidX = XE / 2
MidY = YE / 2
EcartXm = 0.2 * MidX
EcartYm = 0.3 * MidY
EcartXp = 0.8 * MidX
EcartYp = 0.1 * MidY
MXs3d = Abs(EcartXm + EcartXp) / (Xmax - Xmin)
'  *********************************************
'       Tracé en pointillé des contours du
'  domaine de définition dans le plan horizontal
'  *********************************************
If FenetreSurf3D.chkPlanXOY.Value = 1 Then
   pctVueS3D.DrawStyle = 2
   pctVueS3D.Line (MidX - EcartXm, MidY + EcartYm)-(MidX - EcartXp, MidY - EcartYp)
   pctVueS3D.Line -(MidX + EcartXm, MidY - EcartYm)
   pctVueS3D.Line -(MidX + EcartXp, MidY + EcartYp)
   pctVueS3D.Line -(MidX - EcartXm, MidY + EcartYm)
   pctVueS3D.DrawStyle = 0
End If
'  ******************************************
'  Domaine de définition
'  ******************************************
For is3d% = 0 To 2 * NbPasS3D%
   MX%(is3d%) = -10000
   MI%(is3d%) = 10000
Next is3d%
'  ****************************************
'  Tracé
'  ****************************************
On Error GoTo TraiteErreursS3D
pctVueS3D.ForeColor = FenetreSurf3D.pctColorCourbe.BackColor
js3d% = 0
Do While js3d% < NbLiS3D%
   is3d% = 0
   XP = MidX - EcartXm + js3d% * (-EcartXp + EcartXm) / NbLiS3D%
   YP = MidY + EcartYm + js3d% * (-EcartYm - EcartYp) / NbLiS3D%
   Xs3d = Xmin
   Ys3d = Ymin + (Ymax - Ymin) * js3d% / NbLiS3D%
   Direction% = -1
   SigneZ% = 1
   Call PlacerStylo
   If Erreur = True Then Exit Sub
   Call DessinSegment
   ' **********
   ' Bord
   ' **********
   Call PointDeDepart
   Do While is3d% < NbPasS3D%
      DoEvents
      If LanceTrace% = False Or FermeVisuS3D% = True Then Exit Sub
      Direction% = 1
      SigneZ% = 1
      Call PlacerStylo
      Call DessinSegment
      is3d% = is3d% + LongPasS3D%
      Direction% = -1
      Call DessinSegment
      Call PointDeDepart
   Loop
   js3d% = js3d% + 1
Loop
is3d% = 0
XP = MidX - EcartXp
YP = MidY - EcartYp
Xs3d = Xmin
Ys3d = Ymax
js3d% = NbLiS3D%
Do While is3d% < NbPasS3D%
   ' **********
   '  Bord
   ' **********
   Direction% = 1
   SigneZ% = 1
   Call PlacerStylo
   Call DessinSegment
   is3d% = is3d% + LongPasS3D%
Loop
' ***************************************************
' Crosmarie-Perron-Philippine
' Basic Techniques de Programmation, Sybex Paris 1985
' ***************************************************
Exit Sub
TraiteErreursS3D:
    Select Case Err
        Case 5, 6, 11
            Resume Next
        Case Else
            Resume
    End Select
End Sub

Private Sub chkGarder_Click()
   If chkGarder.Value = 0 Then
      Gap3D% = 0
   ElseIf chkGarder.Value = 1 Then
      Gap3D% = 1
   End If
End Sub





Private Sub mnuRetour_Click()
   FermeVisuS3D% = True
   FenetreVisuS3D.Hide
End Sub

Private Sub mnuTraceStop_Click()
   Select Case LanceTrace%
   Case True
      mnuTraceStop.Caption = "&Tracer"
      lblMessage.Caption = ""
      LanceTrace% = False
   Case False
      mnuTraceStop.Caption = "&Arrêter le tracé"
      lblMessage.Caption = "Tracé en cours ..."
      LanceTrace% = True
      Call CalculeS3D
      pctVueS3D.ForeColor = NOIR
      mnuTraceStop.Caption = "&Tracer"
      lblMessage.Caption = ""
      LanceTrace% = False
   End Select
End Sub

Private Sub DessineS3D()
'  ****************************************************
'  DessineS3D
'  ****************************************************
   CX = Int(Xa + 0.5)
   If LigneVisibleFlag% = 1 Then
      pctVueS3D.Line -(CX, CY)
   Else
   pctVueS3D.PSet (CX, CY), pctVueS3D.Point(CX, CY)
   End If
'  ****************************************************
End Sub

Private Sub DessinSegment()
'  **********************************************
'  Dessin segment
'  **********************************************
   XM = XP
   YM = YP
   XX = Xs3d
   YY = Ys3d
   For CMP% = 1 To LongPasS3D%
      is3d% = is3d% + Direction%
      Call ProchainPoint
   Next CMP%
   is3d% = is3d% - LongPasS3D% * Direction%
'  **********************************************
End Sub

Private Sub FonctionS3D()
   X = Xs3d
   Y = Ys3d
   vvar$(1) = LTrim$(Format(X, "0.000000"))
   vvar$(2) = LTrim$(Format(Y, "0.000000"))
   Call variable(ForVar2$, ValVar2$)
   Call Traitement(ValVar2$, SorVar2$)
   Z = CSng(SorVar2$)
   Zs3d = MXs3d * Z
End Sub

Private Sub Form_Activate()
   '--------------------------------------------------
   ' *************************************************
   ' ****       paramètres pour le tracé        ******
   ' ****    dans la zone d'image pctVueS3D     ******
   ' *************************************************
   DES3D% = FenetreVisuS3D.pctVueS3D.ScaleWidth
   '---------------------------------------------------
   FenetreVisuS3D.pctVueS3D.ForeColor = NOIR
   FermeVisuS3D% = False
End Sub

Private Sub LigneCachee()
'  ********************************************
'  Ligne cachée
'  ********************************************
   ks3d% = is3d% - LongPasS3D% * js3d% + NbPasS3D%
   If (CY > MI%(ks3d%)) And (CY <= MX%(ks3d%)) Then
      LigneVisibleFlag% = 0
   Else
      Call LigneVisible
   End If
'  ********************************************
End Sub

Private Sub LigneVisible()
'  ******************************************
'  Ligne visible
'  ******************************************
   LigneVisibleFlag% = 1
   If CY > MX%(ks3d%) Then
      MX%(ks3d%) = CY
   Else
      MI%(ks3d%) = CY
   End If
'  ******************************************
End Sub

Private Sub mnuEnregistrer_Click()
   On Error GoTo TraiteErreursEnrImVisuS3D
   Maths.ctrlCMDialog.DefaultExt = "bmp"
   Maths.ctrlCMDialog.Filter = "Image (*.bmp)|*.bmp"
   Maths.ctrlCMDialog.Flags = &H2&
   Maths.ctrlCMDialog.Action = 2
   SavePicture pctVueS3D.Image, Maths.ctrlCMDialog.filename
   Exit Sub
TraiteErreursEnrImVisuS3D:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuOuvrir_Click()
   On Error GoTo TraiteErreursOuvImVisuS3D
   Maths.ctrlCMDialog.Filter = "Image (*.bmp)|*.bmp"
   Maths.ctrlCMDialog.Flags = &H1000& Or &H800&
   Maths.ctrlCMDialog.CancelError = True
   Maths.ctrlCMDialog.Action = 1
   pctVueS3D.Picture = LoadPicture(Maths.ctrlCMDialog.filename)
   lblFormule.Caption = ""
   Exit Sub
TraiteErreursOuvImVisuS3D:
   Select Case Err
      Case 32755
         ' bouton Annuler
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub mnuQuitter_Click()
   FermeVisuS3D% = True
   FenetreVisuS3D.Hide
   FenetreSurf3D.Hide
End Sub

Private Sub PlacerStylo()
'  ******************************************
'  Placer stylo
'  ******************************************
   Call FonctionS3D
   CX = Int(XP + 0.5)
   CY = Int(YP - SigneZ% * Zs3d + 0.5)
   If CY < 0 Then
      CY = 0
   ElseIf CY > XE - 1 Then
      CY = XE - 1
   End If
   pctVueS3D.PSet (CX, CY), pctVueS3D.Point(CX, CY)
'  ******************************************
End Sub

Private Sub PointDeDepart()
'  ******************************************
'  Point de départ
'  ******************************************
   XP = XM
   YP = YM
   Xs3d = XX
   Ys3d = YY
'  ******************************************
End Sub

Private Sub PreTraiteFonction()
   NombVar% = 2
   var$(1) = "X"
   var$(2) = "Y"
   TextVar2$ = UCase$(TextVar2$)
   Call OteBlancs(TextVar2$, CorVar2$)
   Call Constante(CorVar2$, ForVar2$)
End Sub

Private Sub ProchainPoint()
'  ******************************************
'  Prochain point
'  ******************************************
   Xa = XP + (EcartXm + Direction% * EcartXp) / NbPasS3D%
   YA = YP + (-EcartYm - Direction% * EcartYp) / NbPasS3D%
   Xs3d = Xs3d + (1 + Direction%) / NbPasS3D%
   Ys3d = Ys3d + (1 - Direction%) / NbPasS3D%
   Call FonctionS3D
   CY = Int(YA - SigneZ% * Zs3d + 0.5)
   Call LigneCachee
   If CY < 0 Then
      CY = 0
   ElseIf CY > XE - 1 Then
      CY = XE - 1
   End If
   Call DessineS3D
   XP = Xa
   YP = YA
'  ******************************************
End Sub


Private Sub TraceStop()
Select Case LanceTrace%
   Case True
      mnuTraceStop.Caption = "&Tracer"
      lblMessage.Caption = ""
      LanceTrace% = False
   Case False
      mnuTraceStop.Caption = "&Arrêter le tracé"
      lblMessage.Caption = "Tracé en cours ..."
      LanceTrace% = True
      Call CalculeS3D
      pctVueS3D.ForeColor = NOIR
      mnuTraceStop.Caption = "&Tracer"
      lblMessage.Caption = ""
      LanceTrace% = False
   End Select
End Sub
