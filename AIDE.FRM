VERSION 5.00
Begin VB.Form Aide 
   AutoRedraw      =   -1  'True
   Caption         =   "COURBE     Aide"
   ClientHeight    =   5325
   ClientLeft      =   2025
   ClientTop       =   1755
   ClientWidth     =   4425
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
   ScaleHeight     =   5325
   ScaleWidth      =   4425
   Begin VB.CommandButton CommandeOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "Aide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandeOK_Click()
    Fenetre_Prin.SetFocus
    Aide.Hide
End Sub

Private Sub Form_Load()
    Aide.Cls
    Aide.ForeColor = NOIR
    Aide.Print
    Aide.Print "      ";
    Aide.FontUnderline = -1
    Aide.Print "Pour commencer :"
    Aide.FontUnderline = 0
    Aide.Print
    Aide.Print "   Le menu ";
    Aide.ForeColor = ROUGE
    Aide.Print "Définir... ";
    Aide.ForeColor = NOIR
    Aide.Print "permet de définir une nouvelle"
    Aide.Print "courbe, sur un fond blanc ou par dessus une ou"
    Aide.Print "des courbes déjà tracées ou chargées."
    Aide.Print "   Le menu ";
    Aide.ForeColor = ROUGE
    Aide.Print "Ouvrir... ";
    Aide.ForeColor = NOIR
    Aide.Print "permet de charger une courbe"
    Aide.Print "existante sous forme d'un fichier .BMP"
    Aide.Print
    Aide.Print "      ";
    Aide.FontUnderline = -1
    Aide.Print "Par la suite :"
    Aide.FontUnderline = 0
    Aide.Print
    Aide.Print "   Le bouton ";
    Aide.ForeColor = ROUGE
    Aide.Print "Tracer ";
    Aide.ForeColor = NOIR
    Aide.Print "lance le tracé de la courbe"
    Aide.Print "préalablement définie."
    Aide.Print "   Ce même bouton devient alors"
    Aide.Print "   le bouton ";
    Aide.ForeColor = ROUGE
    Aide.Print "Geler le tacé ";
    Aide.ForeColor = NOIR
    Aide.Print "qui arrête provisoirement "
    Aide.Print "le tracé, tandis que"
    Aide.Print "   le bouton ";
    Aide.ForeColor = ROUGE
    Aide.Print "Arrêter le tracé ";
    Aide.ForeColor = NOIR
    Aide.Print "arrête définitivement"
    Aide.Print "le tracé."
    Aide.Print "   Le menu ";
    Aide.ForeColor = ROUGE
    Aide.Print "Enregistrer... ";
    Aide.ForeColor = NOIR
    Aide.Print "permet de sauvegarder"
    Aide.Print "l'image obtenue sous forme d'un fichier .BMP"
    Aide.Print "   Le menu ";
    Aide.ForeColor = ROUGE
    Aide.Print "Imprimer ";
    Aide.ForeColor = NOIR
    Aide.Print "imprime cette même image"
    Aide.Print "   Le menu ";
    Aide.ForeColor = ROUGE
    Aide.Print "Quitter ";
    Aide.ForeColor = NOIR
    Aide.Print "sort de l'application COURBE"
    Aide.Print "   Le menu ";
    Aide.ForeColor = ROUGE
    Aide.Print "? ";
    Aide.ForeColor = NOIR
    Aide.Print "permet d'accéder à la présente Aide"
    Aide.Print "ainsi qu'aux informations légales"
    Aide.Print
    Aide.Print
    Aide.Print
End Sub

