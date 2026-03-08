VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FenetreDefPts 
   Caption         =   "Définition des points"
   ClientHeight    =   5820
   ClientLeft      =   1950
   ClientTop       =   945
   ClientWidth     =   4710
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
   ScaleHeight     =   5820
   ScaleWidth      =   4710
   Begin MSFlexGridLib.MSFlexGrid GridPts 
      Height          =   3855
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6800
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
   Begin VB.CommandButton CommandeOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtNbPts 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblNbPts 
      Caption         =   "Nombre :"
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
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblPts 
      Caption         =   "Ensemble de points :"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblCoorPts 
      Caption         =   "Coordonnées :"
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
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "FenetreDefPts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandeOK_Click()
   FenetreDefPts.Hide
End Sub

Private Sub Form_Activate()
   '***************************************
   ' actualisation de la grille des points
   '***************************************
   FenetreDefPts.txtNbPts.Text = Format(NbPts%, "0")
   If NbPts% > 0 Then
      For i% = 1 To NbPts%
         FenetreDefPts.gridPts.Row = i%
         FenetreDefPts.gridPts.Col = 1
         FenetreDefPts.gridPts.Text = Format(Xpt(i%), "0.000")
         FenetreDefPts.gridPts.Col = 2
         FenetreDefPts.gridPts.Text = Format(Ypt(i%), "0.000")
      Next i%
   End If
   ' **************************************
End Sub

Private Sub Form_Deactivate()
   On Error GoTo Traite_ErreursDEFPTS
   If NbPts% > 0 Then
      ' ********************************************
      '               ReDims
      ' ********************************************
      Erase Xpt, Ypt
      ReDim Xpt(1 To NbPts%), Ypt(1 To NbPts%)
      ' ********************************************
      ' affectation de leurs valeurs aux coordonnées
      ' ********************************************
      For i% = 1 To NbPts%
         FenetreDefPts.gridPts.Row = i%
         FenetreDefPts.gridPts.Col = 1
         If FenetreDefPts.gridPts.Text = "" Then
            FenetreDefPts.gridPts.Text = "0"
            Xpt(i%) = 0
         Else
            Xpt(i%) = CSng(FenetreDefPts.gridPts.Text)
         End If
         FenetreDefPts.gridPts.Col = 2
         If FenetreDefPts.gridPts.Text = "" Then
            FenetreDefPts.gridPts.Text = "0"
            Ypt(i%) = 0
         Else
            Ypt(i%) = CSng(FenetreDefPts.gridPts.Text)
         End If
      Next i%
      ' ************************************************
      ' activation de certaines cases dans la fenêtre à
      ' partir de laquelle a été ouvert cette fenêtre-ci
      ' ************************************************
      Select Case NumChoisi%
      Case NUMINTEGRALE
        '-----------------------------------------------------
        FenetreDefInteg.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefInteg.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefInteg.PointsTracer.Value = 1                         'Checked
        FenetreDefInteg.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefInteg.PointsRelier.Value = 0                         'Checked
        FenetreDefInteg.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefInteg.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      Case NUMCOUR
        '-----------------------------------------------------
        FenetreDefCourbe.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefCourbe.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefCourbe.PointsTracer.Value = 1                         'Checked
        FenetreDefCourbe.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefCourbe.PointsRelier.Value = 0                         'Checked
        FenetreDefCourbe.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefCourbe.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      Case NUMSOLEQ
        '-----------------------------------------------------
        FenetreDefSolEq.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefSolEq.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefSolEq.PointsTracer.Value = 1                         'Checked
        FenetreDefSolEq.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefSolEq.PointsRelier.Value = 0                         'Checked
        FenetreDefSolEq.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefSolEq.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      Case NUMEQUADIF
        '-----------------------------------------------------
        FenetreDefEqDif.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefEqDif.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefEqDif.PointsTracer.Value = 1                         'Checked
        FenetreDefEqDif.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefEqDif.PointsRelier.Value = 0                         'Checked
        FenetreDefEqDif.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefEqDif.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      Case NUMEQUADIF2
        '-----------------------------------------------------
        FenetreDefEqDif2.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefEqDif2.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefEqDif2.PointsTracer.Value = 1                         'Checked
        FenetreDefEqDif2.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefEqDif2.PointsRelier.Value = 0                         'Checked
        FenetreDefEqDif2.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefEqDif2.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      Case NUMPRIMITIVE
        '-----------------------------------------------------
        FenetreDefPrim.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefPrim.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefPrim.PointsTracer.Value = 1                         'Checked
        FenetreDefPrim.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefPrim.PointsRelier.Value = 0                         'Checked
        FenetreDefPrim.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefPrim.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      Case NUMOINCAR
         '
      Case NUMREGRELIN
         '
      Case NUMINPOLAG
         '
      Case NUMSPLINE
         '
      Case NUMPOLYNOME
        '-----------------------------------------------------
        FenetreDefPolynome.EtiquettePoints.ForeColor = &H80000012     'Texte activé
        FenetreDefPolynome.PointsTracer.ForeColor = &H80000012        'Texte activé
        FenetreDefPolynome.PointsTracer.Value = 1                         'Checked
        FenetreDefPolynome.PointsRelier.ForeColor = &H80000012        'Texte activé
        FenetreDefPolynome.PointsRelier.Value = 0                         'Checked
        FenetreDefPolynome.EtiquetteCoulPoints.ForeColor = &H80000012 'Texte activé
        FenetreDefPolynome.EtiquetteCoulPtsChoisie.ForeColor = &H80000012 'Texte activé
        '-----------------------------------------------------
      End Select
      ' ************************************************
   End If
   Exit Sub
Traite_ErreursDEFPTS:
   Select Case Err
      Case 13
         Message$ = "Erreur dans la frappe des données"
         MsgBox Message$, 48
      Case Else
         MsgBox Error$, 48
   End Select
   Exit Sub
End Sub

Private Sub Form_Load()
   ' **************************************
   ' mise en place de la grille des points
   ' **************************************
   FenetreDefPts.txtNbPts.Text = Format(NbPts%, "0")
   FenetreDefPts.gridPts.Rows = NbPts% + 1
   FenetreDefPts.gridPts.Cols = 3
   FenetreDefPts.gridPts.FixedAlignment(0) = 2
   FenetreDefPts.gridPts.FixedAlignment(1) = 2
   FenetreDefPts.gridPts.FixedAlignment(2) = 2
   FenetreDefPts.gridPts.ColWidth(1) = 1500
   FenetreDefPts.gridPts.ColWidth(2) = 1500
   FenetreDefPts.gridPts.Row = 0
   FenetreDefPts.gridPts.Col = 1
   FenetreDefPts.gridPts.Text = "X"
   FenetreDefPts.gridPts.Col = 2
   FenetreDefPts.gridPts.Text = "Y"
   ' ********************************
End Sub

Private Sub gridPts_KeyPress(KeyAscii As Integer)
   PtsText$ = gridPts.Text
   Select Case KeyAscii
   Case 32 To 168
      PtsCar$ = Chr(KeyAscii)
      PtsText$ = PtsText$ & PtsCar$
      gridPts.Text = PtsText$
   Case 8
      If Len(gridPts.Text) > 0 Then
         PtsText$ = Left$(PtsText$, Len(PtsText$) - 1)
         gridPts.Text = PtsText$
      Else
         Beep
      End If
   End Select
End Sub




Private Sub txtNbPts_Change()
   NbPts% = CInt(txtNbPts.Text)
   gridPts.Rows = NbPts% + 1
   ' ********************************
   ' renumérotation 1ère colonne
   gridPts.Col = 0
   For i% = 1 To NbPts%
      gridPts.Row = i%
      gridPts.Text = Format(i%, "0")
   Next i%
   ' ********************************
End Sub

