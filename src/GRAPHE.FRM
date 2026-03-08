VERSION 5.00
Begin VB.Form Dessin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Graphe"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   7395
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
   MousePointer    =   2  'Cross
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   500
   ScaleMode       =   0  'User
   ScaleWidth      =   500
End
Attribute VB_Name = "Dessin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CompteCar%
Dim Xcar(1 To 100)
Dim Ycar(1 To 100)
Dim car(1 To 100)

Private Sub Form_Activate()
    '----------------------------------------------------
    '---- Effacer éventuellement le dessin précédent ----
    '----------------------------------------------------
    If Gap% = 1 Then
        Dessin.Picture = LoadPicture()
    End If
    '----------------------------------------------------
    ' Donner une valeur non nulle à MX pour éviter
    ' une erreur du type 'Division par zéro'
    ' dans la Function Xreel
    If MX = 0 Then
      MX = 45
    End If
    '----------------------------------------------------
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If DoFlag = False Then
      Select Case KeyAscii
      Case 32 To 168
         CompteCar% = CompteCar% + 1
         Xcar(CompteCar%) = CurrentX
         Ycar(CompteCar%) = CurrentY
         car(CompteCar%) = Chr(KeyAscii)
         Print car(CompteCar%);
      Case 8
         If CompteCar% > 0 Then
            Dessin.PSet (Xcar(CompteCar%), Ycar(CompteCar%)), Dessin.BackColor
            Dessin.ForeColor = Dessin.BackColor
            Print car(CompteCar%);
            Dessin.PSet (Xcar(CompteCar%), Ycar(CompteCar%)), Dessin.BackColor
            Dessin.ForeColor = NOIR
            CompteCar% = CompteCar% - 1
         Else
            Beep
         End If
      End Select
   End If
End Sub

Private Sub Form_Load()
    LX% = Dessin.ScaleWidth
    LY% = Dessin.ScaleHeight
    ' Facteur d'échelle
    FE = LY% / LX%
    If Ech% = 1 Then
        Ymax = FE * (Xmax - Xmin) + Ymin
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
   If DoFlag = False Then
      CompteCar% = 0
      PSet (Xmouse, Ymouse), Dessin.BackColor
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
    Xpoint = Xreel(Xmouse)
    Ypoint = Yreel(Ymouse)
    Fenetre_Val.ForeColor = Fenetre_Val.BackColor
    Fenetre_Val.PSet (0, 4200), Fenetre_Val.BackColor
    Fenetre_Val.Print Space$(50)
    Fenetre_Val.Print Space$(50)
    Fenetre_Val.ForeColor = NOIR
    Fenetre_Val.PSet (0, 3900), Fenetre_Val.BackColor
    Fenetre_Val.Print " Coordonnées du pointeur :"
    Fenetre_Val.PSet (0, 4200), Fenetre_Val.BackColor
    Fenetre_Val.Print " X="; Xpoint
    Fenetre_Val.Print " Y="; Ypoint
End Sub

Private Sub Form_Resize()
    LX% = Dessin.ScaleWidth
    LY% = Dessin.ScaleHeight
    ' Facteur d'échelle
    FE = LY% / LX%
    If Ech% = 1 Then
        AX = Xmax - Xmin
        AY = Ymax - Ymin
        AY0 = AY
        AY = FE * AX
        Ymin = Ymin * AY / AY0
        Ymax = AY + Ymin
    End If
End Sub

Private Function Xreel(Xmouse)
    Xreel = Xmouse / MX + Xmin
End Function


Private Function Yreel(Ymouse)
    Yreel = Ymouse / MY + Ymax
End Function

