Attribute VB_Name = "MATH1"
'                          ****************
'                          *    COURBE    *
'                          ****************
'------------------------ Dims et constantes ---------------------------
Option Base 1
' -----------------------------------
Public DoFlag As Boolean
Public Message$
Public var$(3), vvar$(3)
Public FoNum$(11), kr$(120)
Public X, Y, Z, TETA, R, T
Public Coor%, Expr%, Gap%, Ech%
Public Trac%, Grac%, Vac%, Tric%
Public Trax%, Grax%, Vax%
Public Tran%, Trap%, Rep%, Gap3D%
Public Xpar, Ypar, Xmin, Xmax, Ymin, Ymax, Zmin, Zmax
Public Tmin, Tmax, Xmouse, Ymouse
Public Xvnd, Yvnd, Zvnd ' valeurs numériques données
Public TextVar1$, CorVar1$, ForVar1$, ValVar1$, SorVar1$, ParVar1$
Public TextVar2$, CorVar2$, ForVar2$, ValVar2$, SorVar2$, ParVar2$
Public NombVar%
Public LX%, LY%, FE, AX, AY, RX, RY, MX, MY
Public SE%, SD%, Erreur As Boolean
Public CouleurCou&, CouleurPts&, CouleurHach&, CouleurTan&
Public Const PI = 3.141592653589
' CONVDEGRAD = 180 / PI
Public Const CONVDEGRAD = 57.2957795130968
Public Const M = 0.434294481903
' TAU = nombre d'or = [1+Sqr(5)]/2
' Propriété : UAT = 1/TAU = TAU-1
Public Const TAU = 1.618033989, UAT = 0.618033989
'--------------------- Couleurs ----------------------------------------
Public Const COULEUR_ARRIERE = &H8000000F
Public Const COULEUR_AVANT = &H80000008
Public Const NOIR = &H0&
Public Const BLEU = &HFF0000
Public Const VERT = &HFF00&
Public Const CYAN = &HFFFF00
Public Const ROUGE = &HFF&
Public Const MAGENTA = &HFF00FF
Public Const JAUNE = &HFFFF&
Public Const BLANC = &HFFFFFF
'-----------------------------------------------------------------------
'                          ****************
'                          *    DIVERS    *
'                          ****************
'------------------------ Dims et constantes ---------------------------
Public Const NUMCALC = 0
Public Const NUMBLANC1 = 1
Public Const NUMARITHM = 2
Public Const NUMBLANC2 = 3
Public Const NUMDECFAPRE = 4
Public Const NUMPPCMPGCD = 5
Public Const NUMNOMBPREM = 6
Public Const NUMTRINOMB = 7
Public Const NUMCOMPLEXES = 8
Public Const NUMBLANC3 = 9
Public Const NUMALGEBRE = 10
Public Const NUMBLANC4 = 11
Public Const NUMPOLYNOME = 12
Public Const NUMSOLEQ = 13
Public Const NUMALGLIN = 14
Public Const NUMATRICE = 15
Public Const NUMSYSLIN = 16
Public Const NUMSYSNONLIN = 17
Public Const NUMPROLIN = 18
Public Const NUMBLANC5 = 19
Public Const NUMANALYSE = 20
Public Const NUMBLANC6 = 21
Public Const NUMCOUR = 22
Public Const NUMDERIVE = 23
Public Const NUMEQUADIF = 24
Public Const NUMEQUADIF2 = 25
Public Const NUMPRIMITIVE = 26
Public Const NUMINTEGRALE = 27
Public Const NUMINTERAP = 28
Public Const NUMOINCAR = 29
Public Const NUMREGRELIN = 30
Public Const NUMINPOLAG = 31
Public Const NUMSPLINE = 32
Public Const NUMBLANC7 = 33
Public Const NUMGEOM = 34
Public Const NUMBLANC8 = 35
Public Const NUMTRIANGLE = 36
Public Const NUMPOLYEDRE = 37
Public Const NUMSURF3D = 38
Public NumChoisi%
Public eps          ' epsilon : quantité infinitésimale
'--------------------------------------
' *************************************
' **************  Points **************
' *************************************
' -----------------------------------
' Coordonnées de points :
Public Xpt(), Ypt()
' -----------------------------------
' Nombre de points
Public NbPts%
' -------------------------------------
' *************************************
' *************  SolEq *************
' *************************************
' -----------------------------------
' Bascule :
Public NewtonConverge As Boolean
' FALSE si NewtonRaphson ne converge pas,
' TRUE sinon
' -----------------------------------
Public Xapproche, Xsolution
' -------------------------------------
' ******************************************
' ********  Equation différentielle ********
' ********     du premier ordre     ********
' ******************************************
' -----------------------------------
'
' -----------------------------------
' ******************************************
' ********  Equation différentielle ********
' ********      du second ordre     ********
' ******************************************
' -----------------------------------
'
' -----------------------------------
' *************************************
' *************  Primitive  ***********
' *************************************
' -----------------------------------

' -----------------------------------
' *************************************
' *************  Integrale  ***********
' *************************************
' -----------
' bornes :
Public Xa, Xb
' -----------
' valeur de l'intégrale
Public Integrale
' -------------------------------------
' *************************************
' *************   Dérivée   ***********
' *************************************
' -----------
' variable :
Public nomvar$
' -----------
' expression à dériver :
Public expression$
' -----------
' expression dérivée :
Public exprderivee$
' -------------------------------------
' *************************************************
' ***********  Système non linéaire  **************
' *************************************************
' Nombre d'équations
Public NbEq%
' Matrice
' Mmat() ; déjà déclarée dans ** Matrices **
Public SENL$()
' Vecteurs :
' Uvec(), Vvec() ; déjà déclarés dans ** Matrices **
' Upar(), UApar(); déja déclarés dans ** Moindres carrés **
'ordre :
' OrdreMat% ; déjà déclaré dans ** Matrices **
'--------------------------------------
' *************************************
' *************  Matrices *************
' *************************************
Public Mmat(), M1mat(), M2mat()
Public Pmat(), Wmat()
' Vecteurs :
Public Uvec(), Vvec()
' produit des éléments d'une matrice,
' utilisé pour le calcul du déterminant :
Public pem()
' déterminants :
Public Det, DetMmat
' ordre :
Public OrdreMat%
'--------------------------------------
' Permutations
Public Permut(), SignePermut()
' -------------------------------------
' *************************************
' *************  Polynômes  ***********
' *************************************
' -----------------------------------
' coefficients :
Public Ppol(), P2pol(), PPpol(), Qpol(), Rpol()
' degrés :
Public DegPpol%, DegP2pol%, DegPPpol%, DegQpol%, DegRpol%
' parties réelles et imaginaires
' des racines de Ppol et leurs modules :
Public RacineR(), RacineI(), ModuleRac()
' drapeau signalant l'existence de racines complexes
Public racinecomplexe%
'--------------------------------------
' *************************************
' **********  Moindres carrés *********
' *************************************
' -----------------------------------
' Coordonnées de points :
Public Xmc(), Ymc()
' -----------------------------------
' Paramètres dans une fonction :
Public Upar(), UApar()
' -----------------------------------
' Nombre de paramètres
Public NbPar%
' Nombre de points
Public NbPtsMoinCar%
' -------------------------------------
' *****************************************
' **********  Regression linéaire *********
' *****************************************
' -----------------------------------
' Coordonnées de points :
Public Xrl(), Yrl()
' -------------------------------------
' Paramètres
Public Adroite, Bdroite
' Nombre de points
Public NbPtsReLin%
' -------------------------------------
' Coefficient de corrélation
Public Rrl
' -------------------------------------
' ************************************************
' ****  Polynômes d'interpolation de Lagrange ****
' ************************************************
' -----------------------------------
' Coordonnées de points :
Public Xpl(), Ypl()
' -----------------------------------
' Paramètres dans une fonction :
Public UparI()
' -----------------------------------
' Degré du polynôme de Lagrange
Public DegPoLag%
' Nombre de points
Public NbPtsI%
' -------------------------------------
' Paramètres utilisés dans l'algorithme d'Aitken
Public Fpolag(), Gpolag(), Hpolag()
' -------------------------------------
' ************************************************
' *******************  Spline ********************
' ************************************************
' -----------------------------------
' Coordonnées de points :
Public Xsp(), Ysp()
' -----------------------------------
' Nombre de points
Public NbPtsSp%
' -------------------------------------
' Paramètres utilisés dans le calcul
Public Aspline(), Bspline(), Cspline(), Dspline()
Public Espline(), Fspline(), Gspline(), Hspline()
' -------------------------------------
' *************************************
' ************* Triangles *************
' *************************************
' coordonnées des sommets
Public Xtrian(1 To 3), Ytrian(1 To 3)
' angles (en radians) et côtés
Public AngleTrian(1 To 3), Cote(1 To 3)
' différences de coordonnées des sommets
Public DXtrian(1 To 3), DYtrian(1 To 3)
' 0 si pas d'angle obtus, n° de l'angle obtus sinon
Public Obtus%
' surface du triangle
Public AireTrian
' demi-périmètre
Public DemiPerTrian
' coordonnées du Barycentre
Public Xbar, Ybar
' coordonnées de l'Orthocentre
Public Xort, Yort
' coordonnées du centre du cercle inscrit
Public Xins, Yins
' coordonnées du centre du cercle circonscrit
Public Xcir, Ycir
' coordonnées du centre du cercle d'Euler
Public Xeul, Yeul
' rayon du cercle inscrit
Public RayonInscrit
' rayon du cercle circonscrit
Public RayonCirc
' rayon du cercle d'Euler
Public RayonEuler
' rayon maximum du triangle :
' rayon du cercle circonscrit
' - au triangle (si tous les angles sont aigus) ou
' - au triangle et à son Orthocentre (sinon)
Public RayonMaxTrian
' coordonnées du centre de ce cercle :
Public Xcentrecran, Ycentrecran
' demi-périmètre moins coté i%
Public DemiPerMoinsCote(1 To 3)
' sinus de l'angle i%
Public SinAng(1 To 3)
' cosinus de l'angle i%
Public CosAng(1 To 3)
' coordonnées des pieds des médianes
Public Xmediane(1 To 3), Ymediane(1 To 3)
' longueur des médianes
Public Lmediane(1 To 3)
' coordonnées des pieds des hauteurs
Public Xhauteur(1 To 3), Yhauteur(1 To 3)
' longueur des hauteurs
Public Lhauteur(1 To 3)
' coordonnées des pieds des bissectrices
Public Xbissec(1 To 3), Ybissec(1 To 3)
' longueur des bissectrices
Public Lbissec(1 To 3)
' coordonnées des intersections des médiatrices
' avec le coté du triangle autre que celui
' perpendiculaire à la médiatrice
Public Xmediat(1 To 3), Ymediat(1 To 3)
' longueur des médiatrices
Public Lmediat(1 To 3)
' -------------------------------------
' *************************************
' ************* Polyèdres *************
' *************************************
' Nom du polyèdre
Public NomPolyedre$
' Nombre de sommets
Public NbSom%
' Nombre d'arêtes
Public NbAr%
' Nombre de faces maximum
Public NbFacMax%
' Nombre de faces
Public NbFac%
' Coordonnées de sommets
Public Xsom(), Ysom(), Zsom()
' Arêtes
' 0 : pas d'arête; 1 : arête visible; 2 : arête invisible
Public Arete%()
' Rayon
Public Rayon()
' Rayon maximum
Public RayonMax
' Numéro du point constituant
' le NSième sommet de la face JF
Public NumSomFac%()
' signe +1 ou -1 à appliquer au produit vectoriel V des
' 2 premières arêtes de la face JF pour qu'il soit dirigé
' vers l'extérieur du polyèdre
Public SignePV%()
' nombre de sommets de la face JF
Public NbSomFac%()
' coordonnées du centre de gravité de la face JF
Public Xfac(), Yfac(), Zfac()
' -----------------------------------

Sub Constante(achaine$, bchaine$)
   Rem remplace les constantes par leurs valeurs
   bchaine$ = achaine$
   i3% = 1
   Do
      l% = Len(bchaine$)
      Rem teste la présence d'une fonction numérique et la saute
      Do While l% - i3% > 2
         FoFlag% = 0
         tvar$ = Mid$(bchaine$, i3%, 4)
         For i0% = 1 To 11
            If tvar$ = FoNum$(i0%) Then
               i3% = i3% + 4
               FoFlag% = 1
               Exit For
            End If
         Next i0%
         If FoFlag% = 0 Then Exit Do
      Loop
      Rem teste la présence d'une constante et la remplace
      If l% - i3% > 0 Then
         tvar$ = Mid$(bchaine$, i3%, 2)
         If tvar$ = "PI" Then
            If i3% > 1 Then
               aa3$ = Left$(bchaine$, i3% - 1)
            Else
               aa3$ = ""
            End If
            bb3$ = Mid$(bchaine$, i3% + 2)
            bchaine$ = aa3$ + LTrim$(Format(PI)) + bb3$
            i3% = 0
         End If
      End If
      i3% = i3% + 1
      If i3% > l% Then Exit Do
   Loop
End Sub

Sub DetMat(NombrEntier%)
   Det = 0
   FactN% = Factorielle(NombrEntier%)
   ReDim pem(FactN%)
   For k% = 1 To FactN%
      pem(k%) = SignePermut(k%)
      For s% = 1 To NombrEntier%
         pem(k%) = pem(k%) * M1mat(Permut(k%, s%), s%)
      Next s%
      Det = Det + pem(k%)
   Next k%
End Sub

Sub Evaluation(acar$, resul$)
   Rem cherche les nombres et les opérateurs numériques dans la chaine acar$
   Rem (sans parenthèses), effectue les opérations et renvoie resul$
   Dim operateur$(1 To 100)
   Dim nombre$(1 To 100)
   For iev% = 1 To 15
      nombre$(iev%) = ""
      operateur$(iev%) = ""
   Next iev%
   Erreur = False
   Do
      techaine$ = acar$
      operateur$(1) = ""
      Fin% = 0
      jev% = 0
      iev% = 1
      Do
         Rem cherche nombre$(jev%) à partir de la position iev%
         jev% = jev% + 1
         bon% = 0
         indicchiffre% = 0
         indicpoint% = 0
         c$ = ""
         bcar$ = Mid$(techaine$, iev%, 1)
         If bcar$ = "+" Or bcar$ = "-" Then
            c$ = c$ + bcar$
            iev% = iev% + 1
         End If
         If bcar$ = " " Then
            iev% = iev% + 1
         End If
         Do
            bcar$ = Mid$(techaine$, iev%, 1)
            Select Case bcar$
               Case "0" To "9"
                  indicchiffre% = 1
               Case "E", "e", "D", "d"
                  indicchiffre% = 0
                  d$ = Mid$(techaine$, iev% + 1, 1)
                  If d$ = " " Or d$ = "+" Or d$ = "-" Then
                     c$ = c$ + bcar$
                     iev% = iev% + 1
                     bcar$ = d$
                  End If
               Case ".", ","
                  If indicpoint% = 0 Then
                     indicpoint% = 1
                  Else
                     Erreur = True
                     Message$ = "succession de deux points ou virgules"
                     MsgBox Message$, 16
                     Exit Sub
                  End If
               Case Else
                  If indicchiffre% = 0 Then
                     If bcar$ = "-" Then
                        If Mid$(techaine$, iev% - 1, 1) = "-" Then
                           c$ = Left$(c$, Len(c$) - 1)
                           bcar$ = "+"
                        ElseIf Mid$(techaine$, iev% - 1, 1) = "+" Then
                           c$ = Left$(c$, Len(c$) - 1)
                        Else
                           Erreur = True
                           Message$ = "formule incorrecte !!"
                           MsgBox Message$, 16
                           Exit Sub
                        End If
                     ElseIf bcar$ = "+" Then
                        If Mid$(techaine$, iev% - 1, 1) = "-" Then
                           c$ = Left$(c$, Len(c$) - 1)
                           bcar$ = "-"
                        ElseIf Mid$(techaine$, iev% - 1, 1) = "+" Then
                           c$ = Left$(c$, Len(c$) - 1)
                        Else
                           Erreur = True
                           Message$ = "formule incorrecte !!"
                           MsgBox Message$, 16
                           Exit Sub
                        End If
                     Else
                        Erreur = True
                        Message$ = "formule incorrecte !!"
                        MsgBox Message$, 16
                        Exit Sub
                     End If
                  Else
                     bon% = 1
                  End If
            End Select
            If bon% = 1 Then
               nombre$(jev%) = c$
               Exit Do
            Else
               c$ = c$ + bcar$
               iev% = iev% + 1
            End If
         Loop
         If bcar$ = "" Then
            Fin% = 1
            Exit Do
         End If
         Rem cherche operateur$(jev%) à partir de la position iev%
         If bcar$ = "*" Or bcar$ = "/" Or bcar$ = "+" Or bcar$ = "-" Or bcar$ = "^" Then
            operateur$(jev%) = bcar$
         Else
            Erreur = True
            Message$ = "signe incorrect"
            MsgBox Message$, 16
            Exit Sub
         End If
         iev% = iev% + 1
      Loop
   Loop Until Fin% = 1
   iev% = 1
   Do
      If operateur$(iev%) = "^" Then
         resulop = CSng(nombre$(iev%)) ^ CSng(nombre$(iev% + 1))
         nombre$(iev%) = Format(resulop, "0.000000")
         For jev% = iev% To 14
            operateur$(jev%) = operateur$(jev% + 1)
         Next jev%
         For jev% = iev% + 1 To 14
            nombre$(jev%) = nombre$(jev% + 1)
         Next jev%
      Else
         iev% = iev% + 1
      End If
      If operateur$(iev%) = "" Then Exit Do
   Loop
   iev% = 1
   Do
      If operateur$(iev%) = "*" Then
         resulop = CSng(nombre$(iev%)) * CSng(nombre$(iev% + 1))
         nombre$(iev%) = Format(resulop, "0.000000")
         For jev% = iev% To 14
            operateur$(jev%) = operateur$(jev% + 1)
         Next jev%
         For jev% = iev% + 1 To 14
            nombre$(jev%) = nombre$(jev% + 1)
         Next jev%
      Else
         iev% = iev% + 1
      End If
      If operateur$(iev%) = "" Then Exit Do
   Loop
   iev% = 1
   Do
      If operateur$(iev%) = "/" Then
         resulop = CSng(nombre$(iev%)) / CSng(nombre$(iev% + 1))
         nombre$(iev%) = Format(resulop, "0.000000")
         For jev% = iev% To 14
            operateur$(jev%) = operateur$(jev% + 1)
         Next jev%
         For jev% = iev% + 1 To 14
            nombre$(jev%) = nombre$(jev% + 1)
         Next jev%
      Else
         iev% = iev% + 1
      End If
      If operateur$(iev%) = "" Then Exit Do
   Loop
   iev% = 1
   Do
      If operateur$(iev%) = "-" Then
         resulop = CSng(nombre$(iev%)) - CSng(nombre$(iev% + 1))
         nombre$(iev%) = Format(resulop, "0.000000")
         For jev% = iev% To 14
            operateur$(jev%) = operateur$(jev% + 1)
         Next jev%
         For jev% = iev% + 1 To 14
            nombre$(jev%) = nombre$(jev% + 1)
         Next jev%
      Else
         iev% = iev% + 1
      End If
      If operateur$(iev%) = "" Then Exit Do
   Loop
   iev% = 1
   Do
      If operateur$(iev%) = "+" Then
         resulop = CSng(nombre$(iev%)) + CSng(nombre$(iev% + 1))
         nombre$(iev%) = Format(resulop, "0.000000")
         For jev% = iev% To 14
            operateur$(jev%) = operateur$(jev% + 1)
         Next jev%
         For jev% = iev% + 1 To 14
            nombre$(jev%) = nombre$(jev% + 1)
         Next jev%
      Else
         iev% = iev% + 1
      End If
      If operateur$(iev%) = "" Then Exit Do
   Loop
   resul$ = nombre$(1)
End Sub

Function Factorielle(NombrEntier%)
   If NombrEntier% < 0 Then
      MsgBox "Résultat infini !", 48, "Factorielle"
   ElseIf NombrEntier% = 0 Then
      Factorielle = 1
   ElseIf NombrEntier% = 1 Then
      Factorielle = 1
   Else
      Fact = 1
      For i% = 2 To NombrEntier%
         Fact = Fact * i%
      Next i%
      Factorielle = Fact
   End If
End Function

Sub FoncNum(achaine$, bchaine$)
   coe% = 0
   Rem isole successivement les fonctions numériques dans achaine$,
   Rem l'évalue et renvoie bchaine$ équivalent à achaine$.
   i3% = 1
   bchaine$ = achaine$
   Do
      FoFlag% = 0
      l% = Len(bchaine$)
      bbz$ = Mid$(bchaine$, i3%, 4)
      For i0% = 1 To 11
         If bbz$ = FoNum$(i0%) Then
            FoFlag% = i0%
            Exit For
         End If
      Next i0%
      If FoFlag% <> 0 Then
         j3% = i3%
         Do
            c$ = Mid$(bchaine$, j3%, 1)
            If c$ = "(" Then coe% = coe% + 1
            If c$ = ")" Then
               coe% = coe% - 1
               If coe% = 0 Then Exit Do
            End If
            j3% = j3% + 1
         Loop
         d$ = Mid$(bchaine$, i3% + 4, j3% - i3% - 4)
         aa3$ = Left$(bchaine$, i3% - 1)
         bb3$ = Mid$(bchaine$, j3% + 1)
         Call Traitement(d$, d$)
         Select Case FoFlag%
            Case 1
               d$ = CStr(Abs(CSng(d$)))
            Case 2
               d$ = CStr(Atn(CSng(d$)))
            Case 3
               d$ = CStr(Cos(CSng(d$)))
            Case 4
               d$ = CStr(Exp(CSng(d$)))
            Case 5
               d$ = CStr(Fix(CSng(d$)))
            Case 6
               d$ = CStr(Int(CSng(d$)))
            Case 7
               d$ = CStr(Log(CSng(d$)))
            Case 8
               d$ = CStr(Sgn(CSng(d$)))
            Case 9
               d$ = CStr(Sin(CSng(d$)))
            Case 10
               d$ = CStr(Sqr(CSng(d$)))
            Case 11
               d$ = CStr(Tan(CSng(d$)))
         End Select
         bchaine$ = aa3$ + d$ + bb3$
         i3% = 0
      End If
      i3% = i3% + 1
      If i3% >= l% Then Exit Do
   Loop
End Sub

Sub InvMat01()
'-------------------------------------------------------------------------------
' Inversion d'une matrice carrée
'-------------------------------------------------------------------------------
' Inversion d'une matrice carrée Mmat(i%,j%) d'ordre n%
'
' Méthode absolue, avec calcul du déterminant Detmat
'
' Matrice inverse :  Wmat(i%,j%)
'-------------------------------------------------------------------------------
   n% = OrdreMat%
   If n% < 2 Then
      Beep
      MsgBox "L'ordre de la matrice doit être au moins égal à 2 !", 48, "InvMat01"
      Exit Sub
   End If
   nf% = Factorielle(n%)
   If nf% > 32767 Then
      Beep
      MsgBox "L'ordre de la matrice est trop grand pour cette méthode d'inversion !", 48, "InvMat01"
      Exit Sub
   End If
   ' -----------------------
   ' Initialisations
   ' -----------------------
   eps = 0.0001
   ' -----------------------
   ' DIMs
   ' -----------------------
   ReDim M1mat(1 To n%, 1 To n%), Wmat(1 To n%, 1 To n%)
   ReDim Permut(1 To nf%, 1 To n%), SignePermut(1 To nf%)
   ' -----------------------------------------------
   For i% = 1 To n%
      For j% = 1 To n%
         M1mat(i%, j%) = Mmat(i%, j%)
      Next j%
   Next i%
   ' -----------------------------------------------
   Call Permutations(n%)
   Call DetMat(OrdreMat%)
   If Abs(Det) < eps Then
      Beep
      MsgBox "Le déterminant de la matrice est nul !" & Chr$(10) & "La matrice n'est pas inversible", 48, "InvMat01"
      Exit Sub
   End If
   DetMmat = Det
   nf% = Factorielle(n% - 1)
   ReDim Permut(1 To nf%, 1 To n% - 1), SignePermut(1 To nf%)
   Call Permutations(n% - 1)
   For i% = 1 To n%
      For j% = 1 To n%
         For g% = 1 To n%
            For h% = 1 To n%
               M1mat(g%, h%) = Mmat(g%, h%)
            Next h%
         Next g%
         If j% < n% Then
            For k% = 1 To n%
               For l% = j% To n% - 1
                  M1mat(k%, l%) = M1mat(k%, l% + 1)
               Next l%
            Next k%
         End If
         If i% < n% Then
            For k% = i% To n% - 1
               For l% = 1 To n% - 1
                  M1mat(k%, l%) = M1mat(k% + 1, l%)
               Next l%
            Next k%
         End If
         Call DetMat(OrdreMat% - 1)
         Wmat(j%, i%) = (-1) ^ (i% + j%) * Det / DetMmat
      Next j%
   Next i%
End Sub

Sub InvMat02()
'-------------------------------------------------------------------------------
' Inversion d'une matrice carrée
'-------------------------------------------------------------------------------
' Inversion d'une matrice carrée Mmat(i%,j%) d'ordre n%
'
' Méthode modifiée d'élimination de GAUSS-JORDAN
'
' Matrice inverse :  Wmat(i%,j%)
'-------------------------------------------------------------------------------
   n% = OrdreMat%
   If n% < 2 Then
      Beep
      MsgBox "L'ordre de la matrice doit être au moins égal à 2 !", 48, "InvMat02"
      Exit Sub
   End If
   ReDim M1mat(1 To n%, 1 To n%), Wmat(1 To n%, 1 To n%)
   ' -----------------------------------------------
   For i% = 1 To n%
      For j% = 1 To n%
         M1mat(i%, j%) = Mmat(i%, j%)
      Next j%
   Next i%
   ' -----------------------------------------------
   For i% = 1 To n%
      Wmat(i%, i%) = 1
   Next i%
   ' -----------------------------------------------
   For j% = 1 To n%
      NulsFlag% = 0
      For i% = j% To n%
         If M1mat(i%, j%) <> 0 Then
            NulsFlag% = 1
            Exit For
         End If
      Next i%
      If NulsFlag% = 0 Then
         Beep
         MsgBox "Matrice singulière !", 48, "InvMat02"
         Exit Sub
      End If
      For k% = 1 To n%
         echange = M1mat(j%, k%)
         M1mat(j%, k%) = M1mat(i%, k%)
         M1mat(i%, k%) = echange
         echange = Wmat(j%, k%)
         Wmat(j%, k%) = Wmat(i%, k%)
         Wmat(i%, k%) = echange
      Next k%
      TempVar = 1 / M1mat(j%, j%)
      For k% = 1 To n%
         M1mat(j%, k%) = TempVar * M1mat(j%, k%)
         Wmat(j%, k%) = TempVar * Wmat(j%, k%)
      Next k%
      For l% = 1 To n%
         If l% <> j% Then
            TempVar = -M1mat(l%, j%)
            For k% = 1 To n%
               M1mat(l%, k%) = M1mat(l%, k%) + TempVar * M1mat(j%, k%)
               Wmat(l%, k%) = Wmat(l%, k%) + TempVar * Wmat(j%, k%)
            Next k%
         End If
      Next l%
   Next j%
   ' -----------------------------------------------
End Sub

Sub InvMat03()
'-------------------------------------------------------------------------------
' Inversion d'une matrice carrée
'-------------------------------------------------------------------------------
' Inversion d'une matrice carrée Mmat(i%,j%) d'ordre n%
'
' Méthode itérative passant par le calcul du polynôme caractéristique
'
' Matrice inverse :  Wmat(i%,j%)
'-------------------------------------------------------------------------------
   n% = OrdreMat%
   If n% < 2 Then
      Beep
      MsgBox "L'ordre de la matrice doit être au moins égal à 2 !", 48, "InvMat03"
      Exit Sub
   End If
   eps = 0.0001
   ReDim M1mat(1 To n%, 1 To n%)
   ReDim M3mat(1 To n%, 1 To n%)
   ReDim M4mat(1 To n%, 1 To n%)
   ReDim Wmat(1 To n%, 1 To n%)
   ' -----------------------------------------------
   For i% = 1 To n%
      For j% = 1 To n%
         M1mat(i%, j%) = Mmat(i%, j%)
      Next j%
   Next i%
   ' -----------------------------------------------
   For i% = 1 To n% - 1
      ' **********************************
      ' *        Trace de M1mat          *
      ' **********************************
      Trace = 0
      For j% = 1 To n%
         Trace = Trace + M1mat(j%, j%)
      Next j%
      ' **********************************
      For j% = 1 To n%
         M1mat(j%, j%) = M1mat(j%, j%) - Trace / i%
      Next j%
      For j% = 1 To n%
         For k% = 1 To n%
            For l% = 1 To n%
               M3mat(j%, k%) = M3mat(j%, k%) + Mmat(j%, l%) * M1mat(l%, k%)
            Next l%
         Next k%
      Next j%
      For j% = 1 To n%
         For k% = 1 To n%
            M4mat(j%, k%) = M1mat(j%, k%)
            M1mat(j%, k%) = M3mat(j%, k%)
            M3mat(j%, k%) = 0
         Next k%
      Next j%
   Next i%
   ' **********************************
   ' *******  Trace de M1mat  *********
   ' **********************************
   Trace = 0
   For j% = 1 To n%
      Trace = Trace + M1mat(j%, j%)
   Next j%
   ' **********************************
   SigneDt% = 1
   For i% = 1 To n% - 1
      SigneDt% = -SigneDt%
   Next i%
   DetMmat = SigneDt% * Trace / n%
   If Abs(Trace) < eps Then
      Beep
      MsgBox "Le déterminant de la matrice est nul !" & Chr$(10) & "La matrice n'est pas inversible", 48, "InvMat03"
      Exit Sub
   End If
   ' -------------------------------------------------
   '          Matrice inverse
   ' -------------------------------------------------
   For i% = 1 To n%
      For j% = 1 To n%
         Wmat(i%, j%) = M4mat(i%, j%) * n% / Trace
      Next j%
   Next i%
   ' -----------------------------------------------
End Sub
Sub OrdonnePoints(Nbloc%, Xloc(), Yloc())
   ' ***********************************
   ' ** Ordonne un ensemble de points **
   ' **   suivant les X croissants    **
   ' ***********************************
   'Dim Xloc(1 To Nbloc%)
   'Dim Yloc(1 To Nbloc%)
   For iloc% = 1 To Nbloc% - 1
      Xmloc = Xloc(iloc%)
      Ymloc = Yloc(iloc%)
      For jloc% = iloc% + 1 To Nbloc%
         If Xloc(jloc%) < Xmloc Then
            Xploc = Xmloc
            Yploc = Ymloc
            Xmloc = Xloc(jloc%)
            Ymloc = Yloc(jloc%)
            Xloc(jloc%) = Xploc
            Yloc(jloc%) = Yploc
         End If
      Next jloc%
      Xloc(iloc%) = Xmloc
      Yloc(iloc%) = Ymloc
   Next iloc%
End Sub

Sub OteBlancs(bechaine$, bschaine$)
   Rem enlève les blancs dans bechaine$ et vérifie s'il y a autant de
   Rem parenthèses droites que gauches
   Erreur = False
   l% = Len(bechaine$)
   kr$(l% + 1) = ""
   For i% = 1 To l%
      kr$(i%) = Mid$(bechaine$, i%, 1)
   Next i%
   i% = 1
   compteur1% = 0
   Do
      If kr$(i%) = " " Then
         For j% = i% To l%
            kr$(j%) = kr$(j% + 1)
         Next j%
         l% = l% - 1
      End If
      If kr$(i%) = "(" Then
         compteur1% = compteur1% + 1
      End If
      If kr$(i%) = ")" Then
         compteur1% = compteur1% - 1
      End If
      If compteur1% < 0 Then
         Erreur = True
      End If
      i% = i% + 1
   Loop Until kr$(i%) = "" Or Erreur = True
   acar$ = ""
   For j% = 1 To l%
      acar$ = acar$ + kr$(j%)
   Next j%
   bschaine$ = acar$
   If compteur1% <> 0 Then
      Erreur = True
      If compteur1% > 0 Then
         Message$ = "erreur: plus de parenthèses gauches que droites."
      Else
         Message$ = "erreur: plus de parenthèses droites que gauches."
      End If
      MsgBox Message$
   End If
End Sub

Sub Parametre(achaine$, bchaine$)
   Rem remplace les paramètres par leurs valeurs
   bchaine$ = achaine$
   i3% = 1
   Do
      l% = Len(bchaine$)
      Rem teste la présence d'une fonction numérique et la saute
      Do While l% - i3% > 2
         FoFlag% = 0
         tvar$ = Mid$(bchaine$, i3%, 4)
         For i0% = 1 To 11
            If tvar$ = FoNum$(i0%) Then
               i3% = i3% + 4
               FoFlag% = 1
               Exit For
            End If
         Next i0%
         If FoFlag% = 0 Then Exit Do
      Loop
      Rem teste la présence d'un paramètre et le remplace
      If l% - i3% > 0 Then
         tvar$ = Mid$(bchaine$, i3%, 2)
         If tvar$ = "U(" Then
            If i3% > 1 Then
               aa3$ = Left$(bchaine$, i3% - 1)
            Else
               aa3$ = ""
            End If
            bb3$ = Mid$(bchaine$, i3% + 4)
            ConsInd% = CInt(Mid$(bchaine$, i3% + 2, 1))
            If ConsInd% < 1 Or ConsInd% > NbPar% Then
                MsgBox "indice i de paramètre U(i) hors limites !", 48, "Parametre"
                Exit Sub
            End If
            bchaine$ = aa3$ + LTrim$(Format(Upar(ConsInd%))) + bb3$
            i3% = 0
         End If
      End If
      i3% = i3% + 1
      If i3% > l% Then Exit Do
   Loop
End Sub

Sub Parenthese(pechaine$, sortie$)
   Rem isole successivement la parenthèse centrale dans achaine$,
   Rem l'évalue et renvoie bchaine$ équivalent à achaine$
   i1% = 1
   pschaine$ = pechaine$
   Do
      l% = Len(pschaine$)
      bcar$ = Mid$(pschaine$, i1%, 1)
      If bcar$ = ")" Then
         j1% = i1%
         Do
            c$ = Mid$(pschaine$, j1%, 1)
            If c$ = "(" Then Exit Do
            j1% = j1% - 1
         Loop
         d$ = Mid$(pschaine$, j1% + 1, i1% - j1% - 1)
         Call Evaluation(d$, d1$)
         If Erreur = True Then
            Exit Sub
         End If
         acar$ = Left$(pschaine$, j1% - 1) + d1$ + Mid$(pschaine$, i1% + 1)
         pschaine$ = acar$
         i1% = 0
      End If
      i1% = i1% + 1
      If i1% > l% Then Exit Do
   Loop
   Call Evaluation(pschaine$, sortie$)
End Sub

Sub Permutations(NbEntierPos%)
   For j% = 1 To NbEntierPos%
      Permut(1, j%) = j%
   Next j%
   SignePermut(1) = 1
   i% = 1
   For g% = 2 To NbEntierPos%
      f% = i%
      For k% = 1 To g% - 1
         For i1% = 1 To f%
            i% = i% + 1
            For i2% = 1 To NbEntierPos%
               Permut(i%, i2%) = Permut(i1%, i2%)
            Next i2%
            Aech = Permut(i%, k%)
            Permut(i%, k%) = Permut(i%, g%)
            Permut(i%, g%) = Aech
            SignePermut(i%) = -SignePermut(i1%)
         Next i1%
      Next k%
   Next g%
End Sub

Sub Simpson()
'--------------------------------------
' Méthode de SIMPSON :
' Aire = (y0 + 4 y1 + 2 y2 + 4 y3 +...+ 4 yn-1 + yn) dX/3
' [ où yi = f(Xi) ; n pair]
' Application au calcul de l'intégrale de f(X)
' entre Xa et Xb :
' On prend : dX = (Xb-Xa)/n
'            Xa = X0; y0 = ya = f(Xa);
'            Xb = Xn; yn = yb = f(Xb);
'-------------------------------------------
' Choix de n (pair !!!) :
NbPas% = 100
'-------------------------------------------
dx = (Xb - Xa) / NbPas%
'-------------------------------------------
On Error Resume Next
'--------- Calcul et ajout de ya : -----
X = Xa
vvar$(1) = LTrim$(Format(X, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
If Erreur = True Then
   Exit Sub
End If
Integrale = CSng(SorVar2$)
If Err.Number <> 0 Then
   Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
   MsgBox Message$, 48, "SIMPSON"
   Exit Sub
End If
'--------- Calcul et ajout de yb : -----
X = Xb
vvar$(1) = LTrim$(Format(X, "0.000000"))
Call variable(ForVar2$, ValVar2$)
Call Traitement(ValVar2$, SorVar2$)
If Erreur = True Then
   Exit Sub
End If
Integrale = Integrale + CSng(SorVar2$)
If Err.Number <> 0 Then
   Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
   MsgBox Message$, 48, "SIMPSON"
   Exit Sub
End If
'-------------------------------------------
Integrale = Integrale / 2
X = Xa
k% = -1
l% = 1
'-------------------------------------------
'--------- Calcul et ajout de yi : -------
For i% = 1 To NbPas% - 1
   X = X + dx
   vvar$(1) = LTrim$(Format(X, "0.000000"))
   Call variable(ForVar2$, ValVar2$)
   Call Traitement(ValVar2$, SorVar2$)
   If Erreur = True Then
      Exit Sub
   End If
   k% = -k%
   l% = l% + k%
   Integrale = Integrale + l% * CSng(SorVar2$)
   If Err.Number <> 0 Then
      Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
      MsgBox Message$, 48, "SIMPSON"
      Exit Sub
   End If
Next i%
On Error GoTo 0
'--------- Calcul final : ------------------
Integrale = Integrale * dx * 2 / 3
End Sub

Sub Traitement(acar$, sortie$)
   Call FoncNum(acar$, sorcar$)
   Call Parenthese(sorcar$, sortie$)
End Sub

Sub TrianMat()
'-------------------------------------------------------------------------------
' Résolution d'un système linéaire par triangulation
'-------------------------------------------------------------------------------
' Système de n% équations linéaires à n% inconnues :
'
' Mmat(1,1)*Uvec(1)+...+Mmat(1,j%)*Uvec(j%)+...+Mmat(1,n%)*Uvec(n%) = Vvec(1)
' Mmat(2,1)*Uvec(1)+...+Mmat(2,j%)*Uvec(j%)+...+Mmat(2,n%)*Uvec(n%) = Vvec(2)
' . . . . . . . . .
' Mmat(i%,1)*Uvec(1)+...Mmat(i%,j%)*Uvec(j%)+...+Mmat(i%,n%)*Uvec(n%) = Vvec(i%)
' . . . . . . . . .
' Mmat(n%,1)*Uvec(1)+...Mmat(n%,j%)*Uvec(j%)+...+Mmat(n%,n%)*Uvec(n%) = Vvec(n%)
'
' Calcul des Uvec(j%)
'-------------------------------------------------------------------------------
' Si le système est indéterminé,
' met Erreur = True et sort
' ---------------
' Initialisations
' ---------------
Erreur = False
eps = 0.0001
n% = OrdreMat%
' ---------------
For i% = 1 To n% - 1
   elemax = Mmat(i%, i%)
   d% = i%
   For j% = i% + 1 To n%
      If elemax < Mmat(j%, i%) Then
         elemax = Mmat(j%, i%)
         d% = j%
      End If
   Next j%
   If Abs(elemax) < eps Then
      Erreur = True
      ' Système indéterminé
      Exit Sub
   End If
   elevec = Vvec(i%)
   Vvec(i%) = Vvec(d%)
   Vvec(d%) = elevec
   For k% = i% To n%
      elem = Mmat(i%, k%)
      Mmat(i%, k%) = Mmat(d%, k%)
      Mmat(d%, k%) = elem
   Next k%
   For j% = i% + 1 To n%
      Pivot = Mmat(j%, i%)
      For k% = i% To n%
         Mmat(j%, k%) = Mmat(j%, k%) - Mmat(i%, k%) * Pivot / elemax
      Next k%
      Vvec(j%) = Vvec(j%) - Vvec(i%) * Pivot / elemax
   Next j%
Next i%
If Abs(Mmat(n%, n%)) < eps Then
   Erreur = True
   ' Système indéterminé
   Exit Sub
End If
Uvec(n%) = Vvec(n%) / Mmat(n%, n%)
For i% = n% - 1 To 1 Step -1
   elem = 0
   For j% = i% + 1 To n%
      elem = elem + Mmat(i%, j%) * Uvec(j%)
   Next j%
   Uvec(i%) = (Vvec(i%) - elem) / Mmat(i%, i%)
Next i%
End Sub

Sub variable(achaine$, bchaine$)
   Rem remplace les variables par leurs valeurs
   bchaine$ = achaine$
   For i% = 1 To NombVar%
      If var$(i%) <> "" Then
         lv% = Len(var$(i%))
         i3% = 1
         Do
            l% = Len(bchaine$)
            Rem teste la présence d'une fonction numérique et la saute
            Do While l% - i3% > 2
               FoFlag% = 0
               tvar$ = Mid$(bchaine$, i3%, 4)
               For i0% = 1 To 11
                  If tvar$ = FoNum$(i0%) Then
                     i3% = i3% + 4
                     FoFlag% = 1
                     Exit For
                  End If
               Next i0%
               If FoFlag% = 0 Then Exit Do
            Loop
            Rem teste la présence d'une variable et la remplace
            If i3% <= l% - lv% + 1 Then
               bbz$ = Mid$(bchaine$, i3%, lv%)
               If bbz$ = var$(i%) Then
                  If i3% > 1 Then
                     aa3$ = Left$(bchaine$, i3% - 1)
                  Else
                     aa3$ = ""
                  End If
                  bb3$ = Mid$(bchaine$, i3% + lv%)
                  bchaine$ = aa3$ + vvar$(i%) + bb3$
                  i3% = 0
               End If
            End If
            i3% = i3% + 1
            If i3% > l% Then Exit Do
         Loop
      End If
   Next i%
End Sub


Public Sub ZerosPolBairstow()
' -------------------------------------------------------------
' Recherche des racines d'un polynôme
' par la méthode de BAIRSTOW
' David BLUM
' 12/09/1996
' -------------------------------------------------------------
' Recherche les racines du polynôme Ppol de degré DegPpol%
' S'il y a erreur ou non convergence, fait Erreur%=1 et sort;
' Sinon remplit ModuleRac(), RacineR() et RacineI()
' avec les modules et les parties réelles
' et imaginaires des racines triées par ordre de
' module décroissant.
' -------------------------------------------------------------
' Polynôme Ppol(x) d'ordre n :
' Ppol(x) = a(n) * x^n + a(n-1) * x^(n-1) + ... + a(1) * x + a(0) "
' -------------------------------------------------------------
' Polynôme aBairpol(x) d'ordre n :
' aBairpol(x) = a(0) * x^n + a(1) * x^(n-1) + ... + a(n-1) * x + a(n) "
' ---------------------------------------------------------
' Initialisation
Erreur = False
' ---------------------------------------------------------
' Degré suffisant ?
If DegPpol% < 1 Then
   Erreur = True
   Message$ = " Polynôme de degré inférieur à 1 ! "
   MsgBox Message$, 16, "ZerosPolBairstow"
   Exit Sub
End If
' ---------------------------------------------------------
' Ppol(n) <> 0 ?
' ---------------------------------------------------------
If Abs(Ppol(DegPpol%)) < eps Then
   Erreur = True
   Message$ = " Polynôme de degré inférieur à " & DegPpol% & " ! "
   MsgBox Message$, 16, "ZerosPolBairstow"
   Exit Sub
End If
' ---------------------------------------------------------
' Dims
' ---------------------------------------------------------
ReDim aBairPol(0 To DegPpol%)
ReDim bBairPol(0 To DegPpol%)
ReDim cBairPol(0 To DegPpol%)
ReDim qBairPol(0 To DegPpol%)
ReDim RacineR(1 To DegPpol%)
ReDim RacineI(1 To DegPpol%)
ReDim ModuleRac(1 To DegPpol%)
For iloc% = 1 To DegPpol%
   RacineR(iloc%) = 0
   RacineI(iloc%) = 0
   ModuleRac(iloc%) = 0
Next iloc%
' parties réelle et imaginaire des 2 zéros
' de  x² + p1 * x + q1   :
Dim RacineRLoc(1 To 2)
Dim RacineILoc(1 To 2)
For iloc% = 1 To 2
   RacineRLoc(iloc%) = 0
   RacineILoc(iloc%) = 0
Next iloc%
' ---------------------------------------------------------
' Initialisation de aBairpol(x) et qBairpol(x):
' qBairpol(x) = aBairpol(x)
' ---------------------------------------------------------
For iBair% = 0 To DegPpol%
   aBairPol(iBair%) = Ppol(DegPpol% - iBair%)
   qBairPol(iBair%) = aBairPol(iBair%)
Next iBair%
' ---------------------------------------------------------
' Initialisations
' ----------------------------------------------------------------------
' Drapeau signalant l'existence de racines complexes
racinecomplexe% = 0
' ----------------------------------------------------------------------
' Numéro de la racine
numrac% = 0
' ----------------------------------------------------------------------
' Degré de qBairpol(x) :
DegQBairpol% = DegPpol%
' ---------------------------------------------------------
' Précision de convergence
epsBair = 0.0001
' ---------------------------------------------------------
' Nombre de cycles à effectuer
numcyclesmax% = 100
' ---------------------------------------------------------
If DegPpol% > 2 Then
' ---------------------------------------------------------
' Début des itérations :
   Do
      pBair0 = 0
      qBair0 = 0
      numcycle% = 0
      Do
         numcycle% = numcycle% + 1
         ' ---------------------------------------------------------
         ' Détermination des bBairPol(iBair%) :
         bBairPol(0) = qBairPol(0)
         bBairPol(1) = qBairPol(1) - pBair0 * bBairPol(0)
         For iBair% = 2 To DegQBairpol%
            bBairPol(iBair%) = qBairPol(iBair%) - pBair0 * bBairPol(iBair% - 1) - qBair0 * bBairPol(iBair% - 2)
         Next iBair%
         ' ---------------------------------------------------------
         ' Détermination des cBairPol(iBair%) :
         cBairPol(0) = bBairPol(0)
         cBairPol(1) = bBairPol(1) - pBair0 * cBairPol(0)
         If DegQBairpol% > 3 Then
            For iBair% = 2 To DegQBairpol% - 2
               cBairPol(iBair%) = bBairPol(iBair%) - pBair0 * cBairPol(iBair% - 1) - qBair0 * cBairPol(iBair% - 2)
            Next iBair%
         End If
         cBairPol(DegQBairpol% - 1) = -pBair0 * cBairPol(DegQBairpol% - 2) - qBair0 * cBairPol(DegQBairpol% - 3)
         ' ---------------------------------------------------------
         ' Détermination de pBair1 et qBair1 :
         DBairPol = cBairPol(DegQBairpol% - 2) * cBairPol(DegQBairpol% - 2) - cBairPol(DegQBairpol% - 1) * cBairPol(DegQBairpol% - 3)
         UBairPol = bBairPol(DegQBairpol% - 1) * cBairPol(DegQBairpol% - 2) - bBairPol(DegQBairpol%) * cBairPol(DegQBairpol% - 3)
         VBairPol = bBairPol(DegQBairpol%) * cBairPol(DegQBairpol% - 2) - bBairPol(DegQBairpol% - 1) * cBairPol(DegQBairpol% - 1)
         If Abs(DBairPol) < epsBair Then
            pBair1 = 1
            qBair1 = -1
         Else
            pBair1 = pBair0 + UBairPol / DBairPol
            qBair1 = qBair0 + VBairPol / DBairPol
         End If
         ' ---------------------------------------------------------
         ' Test de convergence
         difp = Abs(pBair1 - pBair0)
         difq = Abs(qBair1 - qBair0)
         If difp < epsBair And difq < epsBair Then
            Exit Do
         End If
         ' ---------------------------------------------------------
         ' Test de nombre de cycles maximum
         If numcycle% = numcyclesmax% Then
            Erreur = True
            Message$ = " Pas de convergence après " + Format(numcyclesmax%, "0") + " cycles. "
            MsgBox Message$, 16, "ZerosPolBairstow"
            Exit Sub
         End If
         ' ---------------------------------------------------------
         ' mise à jour de pBair0 et qBair0 pour la boucle suivante
         pBair0 = pBair1
         qBair0 = qBair1
      Loop
      ' ---------------------------------------------------------
      ' Détermination des zéros de x^2 + pBair1 * x + qBair1 :
      ' 1) Discriminant :
      DisBair = pBair1 * pBair1 - 4 * qBair1
      ' 2) Racines  :
      If DisBair > 0 Then
         SDisBair = Sqr(DisBair)
         RacineRLoc(1) = (-pBair1 - SDisBair) / 2
         RacineILoc(1) = 0
         RacineRLoc(2) = (-pBair1 + SDisBair) / 2
         RacineILoc(2) = 0
      ElseIf DisBair = 0 Then
         RacineRLoc(1) = -pBair1 / 2
         RacineILoc(1) = 0
         RacineRLoc(2) = -pBair1 / 2
         RacineILoc(2) = 0
      Else
         racinecomplexe% = 1
         SDisBair = Sqr(-DisBair)
         RacineRLoc(1) = -pBair1 / 2
         RacineILoc(1) = -SDisBair / 2
         RacineRLoc(2) = -pBair1 / 2
         RacineILoc(2) = SDisBair / 2
      End If
      ' 3) Classement et enregistrement des racines  :
      For numractrouv% = 1 To 2
         RacineR0 = RacineRLoc(numractrouv%)
         RacineI0 = RacineILoc(numractrouv%)
         ModuleRac0 = Sqr(RacineR0 * RacineR0 + RacineI0 * RacineI0)
         numrac% = numrac% + 1
         ' ******* mise en ordre selon ModuleRac décroissant *******
         If numrac% > 1 Then
            For iloc% = numrac% - 1 To 1 Step -1
            If ModuleRac0 > ModuleRac(iloc%) Then
               RacineR(iloc% + 1) = RacineR(iloc%)
               RacineI(iloc% + 1) = RacineI(iloc%)
               ModuleRac(iloc% + 1) = ModuleRac(iloc%)
               If iloc% = 1 Then
                  RacineR(iloc%) = RacineR0
                  RacineI(iloc%) = RacineI0
                  ModuleRac(iloc%) = ModuleRac0
               End If
            Else
               RacineR(iloc% + 1) = RacineR0
               RacineI(iloc% + 1) = RacineI0
               ModuleRac(iloc% + 1) = ModuleRac0
               Exit For
            End If
            Next iloc%
         Else
            RacineR(1) = RacineR0
            RacineI(1) = RacineI0
            ModuleRac(1) = ModuleRac0
         End If
         ' ****************************************************
      Next numractrouv%
      ' ---------------------------------------------------------
      ' Nouveau degré de Q(x) :
      DegQBairpol% = DegQBairpol% - 2
      ' ---------------------------------------------------------
      ' Calcul des coefficients de Q(x) :
      For iBair% = 0 To DegQBairpol%
         qBairPol(iBair%) = bBairPol(iBair%)
      Next iBair%
      ' ---------------------------------------------------------
      ' Test de fin d'itérations
      If DegQBairpol% < 3 Then
         Exit Do
      End If
      ' ---------------------------------------------------------
      ' Nouvelle boucle
   Loop
End If
' ---------------------------------------------------------
' Détermination des zéros de Q(x) résiduel :
If DegQBairpol% = 1 Then
   ' Polynôme de degré 1 :
   RacineRLoc(1) = -qBairPol(1) / qBairPol(0)
   RacineILoc(1) = 0
Else
   ' Polynôme de degré 2 :
   ' 1) Discriminant :
   DisBair = qBairPol(1) * qBairPol(1) - 4 * qBairPol(0) * qBairPol(2)
   ' 2) Racines  :
   If DisBair > 0 Then
      SDisBair = Sqr(DisBair)
      RacineRLoc(1) = (-qBairPol(1) - SDisBair) / 2 / qBairPol(0)
      RacineILoc(1) = 0
      RacineRLoc(2) = (-qBairPol(1) + SDisBair) / 2 / qBairPol(0)
      RacineILoc(2) = 0
   ElseIf DisBair = 0 Then
      RacineRLoc(1) = -qBairPol(1) / 2 / qBairPol(0)
      RacineILoc(1) = 0
      RacineRLoc(2) = -qBairPol(1) / 2 / qBairPol(0)
      RacineILoc(2) = 0
   Else
      racinecomplexe% = 1
      SDisBair = Sqr(-DisBair)
      RacineRLoc(1) = -qBairPol(1) / 2 / qBairPol(0)
      RacineILoc(1) = -SDisBair / 2 / qBairPol(0)
      RacineRLoc(2) = -qBairPol(1) / 2 / qBairPol(0)
      RacineILoc(2) = SDisBair / 2 / qBairPol(0)
   End If
End If
' Classement et enregistrement des zéros de Q(x) résiduel  :
For numractrouv% = 1 To DegQBairpol%
   RacineR0 = RacineRLoc(numractrouv%)
   RacineI0 = RacineILoc(numractrouv%)
   ModuleRac0 = Sqr(RacineR0 * RacineR0 + RacineI0 * RacineI0)
   numrac% = numrac% + 1
   ' ******* mise en ordre selon ModuleRac décroissant *******
   If numrac% > 1 Then
      For iloc% = numrac% - 1 To 1 Step -1
      If ModuleRac0 > ModuleRac(iloc%) Then
         RacineR(iloc% + 1) = RacineR(iloc%)
         RacineI(iloc% + 1) = RacineI(iloc%)
         ModuleRac(iloc% + 1) = ModuleRac(iloc%)
         If iloc% = 1 Then
            RacineR(iloc%) = RacineR0
            RacineI(iloc%) = RacineI0
            ModuleRac(iloc%) = ModuleRac0
         End If
      Else
         RacineR(iloc% + 1) = RacineR0
         RacineI(iloc% + 1) = RacineI0
         ModuleRac(iloc% + 1) = ModuleRac0
         Exit For
      End If
      Next iloc%
   Else
      RacineR(1) = RacineR0
      RacineI(1) = RacineI0
      ModuleRac(1) = ModuleRac0
   End If
   ' ****************************************************
Next numractrouv%
' ---------------------------------------------------------
End Sub

Public Sub NewtonRaphson()
' --------------------------------------------
' SOLUTIONS DE L'EQUATION F(X)=0
' Recherche des solutions de l'équation F(X)=0
' par la méthode de NEWTON-RAPHSON :
' On utilise la formule de récursivité :
' Xi+1 = Xi - F(Xi) / F'(Xi)
' où F'(X) est la dérivée de F(X).
' --------------------------------------------
NewtonConverge = False
X0 = Xapproche
NbIter% = 100
eps = 0.000001
i% = 0
On Error Resume Next
Do
   ' ************ F(X) ***************
   vvar$(1) = LTrim$(Format(X0, "0.000000"))
   Call variable(ForVar2$, ValVar2$)
   Call Traitement(ValVar2$, SorVar2$)
   Y0 = CSng(SorVar2$)
   ' ************ F(X+dX) ***************
   X0plus = X0 + eps
   vvar$(1) = LTrim$(Format(X0plus, "0.000000"))
   Call variable(ForVar2$, ValVar2$)
   Call Traitement(ValVar2$, SorVar2$)
   Y0plus = CSng(SorVar2$)
   ' ************************************
   If Err.Number <> 0 Then
      Message$ = "Erreur n° " & Str(Err.Number) & " provenant de " & Err.Source & Chr(13) & Err.Description
      MsgBox Message$, 48, "SIMPSON"
      Exit Sub
   End If
   ' ************************************
   Derivee = (Y0plus - Y0) / eps
   If Abs(Derivee) > eps Then
      X1 = X0 - Y0 / Derivee
   End If
   If Abs(X1 - X0) < eps Then
      Exit Do
   End If
   i% = i% + 1
   X0 = X1
Loop Until i% = NbIter%
On Error GoTo 0
Xsolution = X1
If i% = NbIter% Then
   NewtonConverge = False
   Message$ = " Le nombre maximum N =" & Format(NbIter%, "0")
   Message$ = Message$ & " d'itérations a été atteint avant la convergence."
   Message$ = Message$ & Chr$(13)
   Message$ = Message$ & "A ce stade, la solution est :"
Else
   NewtonConverge = True
   Message$ = Message$ & Chr$(13)
   Message$ = " Convergence après " & Format(i%, "0") & " itérations. " & Chr$(13) & " Solution :"
End If
End Sub


Public Sub DeriveSomme(exprsomme$, varsomme$, exprsommederivee$)
   '***********************************************
   ' Essaie d'appliquer : (u+v)' = u' + v' :
   ' Examine si on peut découper exprsomme$ en
   ' 2 morceaux séparés par un opérateur + ou - :
   ' - si oui renvoie DeriveProduit du 1er morceau
   ' concaténé par l'opérateur + ou - à
   ' DeriveSomme du 2ème morceau;
   ' - si non renvoie DeriveProduit du 2ème morceau
   ' Le renvoi est fait sous forme de exprsommederivee$
   '***********************************************
   ' ---------------------------------------
   ' Initialisations
   ' ---------------------------------------
   chaine$ = exprsomme$
   varloc$ = varsomme$
   exprsommederivee$ = ""
   longchaine% = Len(chaine$)
   ich% = 1
   idebut% = 1
   ifin% = 1
   ' ---------------------------------------------
   ' Balayage de l'expression
   ' ---------------------------------------------
   Do
      ' ------------------------------------------
      ' examine si l'expression est découpée en
      ' morceaux séparés par des opérateurs + ou -
      ' ------------------------------------------
      car$ = Mid$(chaine$, ich%, 1)
      If car$ = "(" Then
         ' ----------------------------------
         ' dans un premier temps, saute les
         ' expressions entre parenthèses
         ' ----------------------------------
         compteur% = 1
         Do
            ich% = ich% + 1
            car$ = Mid$(chaine$, ich%, 1)
            If car$ = ")" Then
               compteur% = compteur% - 1
            ElseIf car$ = "(" Then
               compteur% = compteur% + 1
            End If
            If compteur% = 0 Then
               Exit Do
            End If
         Loop
      ElseIf car$ = "+" Or car$ = "-" Then
         ' ----------------------------------
         ' isole les morceaux séparés par des
         ' opérateurs + ou -, les dérive,
         ' concatène les dérivées et renvoie
         ' le résultat
         ' ----------------------------------
         ifin% = ich% - 1
         If ifin% > 0 Then
            ' ---------------------------------------
            ' si ifin%=0, c'est que le premier
            ' caractère est + ou -, il s'agit alors
            ' d'un signe et non d'un opérateur
            ' ---------------------------------------
            idebut% = 1
            usomme$ = Mid(chaine$, idebut%, ifin%)
            Call DeriveProduit(usomme$, varloc$, usommederive$)
            idebut% = ich% + 1
            ifin% = longchaine%
            vsomme$ = Mid(chaine$, idebut%, ifin%)
            Call DeriveSomme(vsomme$, varloc$, vsommederive$)
            ' ------------------------------------
            ' examine les cas pouvant conduire
            ' à une simplification de l'écriture :
            ' (u+v)' = u'+ v'
            ' ------------------------------------
            If usommederive$ = "0" And vsommederive$ = "0" Then
               ' u et v sont des nombres ou des constantes
               exprsommederivee$ = "0"
            ElseIf usommederive$ = "0" Then
               ' u est un nombre ou une constante
               ' (u+v)' = v'
               exprsommederivee$ = vsommederive$
            ElseIf vsommederive$ = "0" Then
               ' v est un nombre ou une constante
               ' (u+v)' = u'
                  exprsommederivee$ = usommederive$
            Else
               ' u et v contiennent la variable
               If usommederive$ = "1" And vsommederive$ = "1" Then
                  ' u et v sont la variable
                  ' (u+v)' = 2*u
                  exprsommederivee$ = "2*" & usomme$
               Else
                  ' cas général :
                  ' (u+v)' = u'+ v'
                  exprsommederivee$ = usommederive$ & car$ & vsommederive$
               End If
            End If
            Exit Sub
         End If
      End If
      ich% = ich% + 1
      If ich% > longchaine% Then
         ' ------------------------------------------
         ' l'expression est en 1 seul morceau
         ' ------------------------------------------
         idebut% = 1
         ifin% = longchaine%
         usomme$ = Mid(chaine$, idebut%, ifin%)
         Call DeriveProduit(usomme$, varloc$, usommederive$)
         exprsommederivee$ = usommederive$
         Exit Sub
      End If
   Loop
End Sub

Public Sub DeriveProduit(exprprod$, varprod$, exprprodderivee$)
   '***********************************************
   ' Essaie d'appliquer : (uv)' = u'v + uv'
   ' ou (u/v)' = u'/v - uv'/v² :
   ' Examine si on peut découper exprprod$ en
   ' 2 morceaux u et v séparés par un opérateur
   ' * ou / :
   ' - si oui renvoie une des expressions ci-dessus
   ' avec u' = DeriveFonctionPar(u)
   ' et v' = DeriveFonctionPar(v) ;
   ' - si non renvoie u' = DeriveFonctionPar(u)
   ' avec u = exprprod$
   '***********************************************
   ' ---------------------------------------
   ' Initialisations
   ' ---------------------------------------
   chaine$ = exprprod$
   varloc$ = varprod$
   exprprodderivee$ = ""
   longchaine% = Len(chaine$)
   ich% = 1
   idebut% = 1
   ifin% = 1
   ' ---------------------------------------------
   ' Balayage de l'expression
   ' ---------------------------------------------
   Do
      ' ------------------------------------------
      ' examine si l'expression est découpée en
      ' morceaux séparés par des opérateurs * ou /
      ' ------------------------------------------
      car$ = Mid$(chaine$, ich%, 1)
      If car$ = "(" Then
         ' ----------------------------------
         ' dans un premier temps, saute les
         ' expressions entre parenthèses
         ' ----------------------------------
         compteur% = 1
         Do
            ich% = ich% + 1
            car$ = Mid$(chaine$, ich%, 1)
            If car$ = ")" Then
               compteur% = compteur% - 1
            ElseIf car$ = "(" Then
               compteur% = compteur% + 1
            End If
            If compteur% = 0 Then
               Exit Do
            End If
         Loop
      ElseIf car$ = "*" Then
         ' ----------------------------------
         ' isole les morceaux séparés par des
         ' l'opérateur * , les dérive,
         ' applique (uv)'
         ' et renvoie le résultat
         ' ----------------------------------
         idebut% = 1
         ifin% = ich% - 1
         uproduit$ = Mid(chaine$, idebut%, ifin%)
         Call DeriveFonctionPar(uproduit$, varloc$, uproduitderive$)
         idebut% = ich% + 1
         ifin% = longchaine%
         vproduit$ = Mid(chaine$, idebut%, ifin%)
         Call DeriveProduit(vproduit$, varloc$, vproduitderive$)
         ' ------------------------------------
         ' examine les cas pouvant conduire
         ' à une simplification de l'écriture :
         ' (uv)' = u'v + uv'
         ' ------------------------------------
         If uproduitderive$ = "0" And vproduitderive$ = "0" Then
            ' u et v sont des nombres ou des constantes
            exprprodderivee$ = "0"
         ElseIf uproduitderive$ = "0" Then
            ' u est un nombre ou une constante
            If vproduitderive$ = "1" Then
               ' v est la variable
               ' ( plus éventuellement une constante)
               ' (uv)' = u
               exprprodderivee$ = uproduit$
            Else
               ' (uv)' = uv'
               exprprodderivee$ = uproduit$ & "*(" & vproduitderive$ & ")"
            End If
         ElseIf vproduitderive$ = "0" Then
            ' v est un nombre ou une constante
            If uproduitderive$ = "1" Then
               ' u est la variable
               ' ( plus éventuellement une constante)
               ' (uv)' = v
               exprprodderivee$ = vproduit$
            Else
               ' (uv)' = u'v
               exprprodderivee$ = uproduitderive$ & "*" & vproduit$
            End If
         Else
            ' u et v contiennent la variable
            If uproduitderive$ = "1" And vproduitderive$ = "1" Then
               ' u et v sont la variable
               ' ( plus éventuellement une constante)
               ' (uv)' = u + v
               exprprodderivee$ = uproduit$ & "+" & vproduit$
            ElseIf uproduitderive$ = "1" Then
               ' u est la variable
               ' ( plus éventuellement une constante)
               ' (uv)' = v + uv'
               exprprodderivee$ = vproduit$ & "+" & uproduit$ & "*(" & vproduitderive$ & ")"
            ElseIf vproduitderive$ = "1" Then
               ' v est la variable
               ' ( plus éventuellement une constante)
               ' (uv)' = u'v + u
               exprprodderivee$ = uproduitderive$ & "*" & vproduit$ & "+" & uproduit$
            Else
               ' cas général :
               ' (uv)' = u'v + uv'
               exprprodderivee$ = uproduitderive$ & "*" & vproduit$ & "+" & uproduit$ & "*(" & vproduitderive$ & ")"
            End If
         End If
         Exit Sub
      ElseIf car$ = "/" Then
         ' ----------------------------------
         ' isole les morceaux séparés par
         ' l'opérateur /, les dérive,
         ' applique (u/v)'
         ' et renvoie le résultat
         ' ----------------------------------
         idebut% = 1
         ifin% = ich% - 1
         uproduit$ = Mid(chaine$, idebut%, ifin%)
         Call DeriveFonctionPar(uproduit$, varloc$, uproduitderive$)
         idebut% = ich% + 1
         ifin% = longchaine%
         vproduit$ = Mid(chaine$, idebut%, ifin%)
         Call DeriveProduit(vproduit$, varloc$, vproduitderive$)
         ' ------------------------------------
         ' examine les cas pouvant conduire
         ' à une simplification de l'écriture :
         ' (u/v)' = u'/v - uv'/v²
         ' ------------------------------------
         If uproduitderive$ = "0" And vproduitderive$ = "0" Then
            ' u et v sont des nombres ou des constantes
            exprprodderivee$ = "0"
         ElseIf uproduitderive$ = "0" Then
            ' u est un nombre ou une constante
            If vproduitderive$ = "1" Then
               ' v est la variable
               ' ( plus éventuellement une constante)
               ' (u/v)' = - u/v²
               exprprodderivee$ = "-" & uproduit$ & "/" & vproduit$ & "^2"
            Else
               ' (u/v)' = -uv'/v²
               exprprodderivee$ = "-" & uproduit$ & "*" & vproduitderive$ & "/(" & vproduit$ & ")^2"
            End If
         ElseIf vproduitderive$ = "0" Then
            ' v est un nombre ou une constante
            If uproduitderive$ = "1" Then
               ' u est la variable
               ' ( plus éventuellement une constante)
               ' (u/v)' = 1/v
               exprprodderivee$ = "1/" & vproduit$
            Else
               ' (u/v)' = u'/v
               exprprodderivee$ = uproduitderive$ & "/" & vproduit$
            End If
         Else
            ' u et v contiennent la variable
            If uproduitderive$ = "1" And vproduitderive$ = "1" Then
               ' u et v sont la variable
               ' ( plus éventuellement une constante)
               ' (u/v)' = 1/v - u/v²
               exprprodderivee$ = "1/" & vproduit$ & "-" & uproduit$ & "/" & vproduit$ & "^2"
            ElseIf uproduitderive$ = "1" Then
               ' u est la variable
               ' ( plus éventuellement une constante)
               ' (u/v)' = 1/v - uv'/v²
               exprprodderivee$ = "1/" & vproduit$ & "-" & uproduit$ & "*" & vproduitderive$ & "/" & vproduit$ & "^2"
            ElseIf vproduitderive$ = "1" Then
               ' v est la variable
               ' ( plus éventuellement une constante)
               ' (u/v)' = u'/v - u/v²
               exprprodderivee$ = uproduitderive$ & "/" & vproduit$ & "-" & uproduit$ & "/" & vproduit$ & "^2"
            Else
               ' Cas général
               ' (u/v)' = u'/v - uv'/v²
               exprprodderivee$ = uproduitderive$ & "/" & vproduit$ & "-" & uproduit$ & "*" & vproduitderive$ & "/" & vproduit$ & "^2"
            End If
         End If
         Exit Sub
      End If
      ich% = ich% + 1
      If ich% > longchaine% Then
         ' ------------------------------------------
         ' l'expression est en 1 seul morceau
         ' ------------------------------------------
         ifin% = longchaine%
         uproduit$ = Mid(chaine$, idebut%, ifin%)
         Call DeriveFonctionPar(uproduit$, varloc$, uproduitderive$)
         exprprodderivee$ = uproduitderive$
         Exit Sub
      End If
   Loop
End Sub

Public Sub DeriveFonctionPar(exprfonpar$, varfonpar$, exprfonparderivee$)
   '****************************************************
   ' 1)- Recherche la présence de l'exponentiation ^
   ' 2)- Recherche u, avant le signe ^, de la
   ' forme : ();F();variable;nombre;
   ' 3)- Si le signe ^ est absent, renvoi u'
   ' 4)- Si le signe ^ est présent, détermine
   ' v après le signe ^ et renvoie :
   ' (u^v)' = u^v*v'*LOG(u) + u^(v-1)*u'*v
   ' Le renvoi est fait sous forme de exprfonparderivee$
   '****************************************************
   ' ---------------------------------------
   ' Initialisations
   ' ---------------------------------------
   Erreur = False
   chaine$ = exprfonpar$
   varloc$ = varfonpar$
   exprfonparderivee$ = ""
   longchaine% = Len(chaine$)
   ich% = 1
   iexpon% = 0
   ' ---------------------------------------------
   ' Premier balayage de l'expression
   ' pour rechercher le caractère ^
   ' ---------------------------------------------
   Do
      ' ------------------------------------------
      ' examine l'expression
      ' ------------------------------------------
      car$ = Mid$(chaine$, ich%, 1)
      If car$ = "(" Then
         ' ----------------------------------
         ' dans un premier temps, saute les
         ' expressions entre parenthèses
         ' ----------------------------------
         compteur% = 1
         Do
            ich% = ich% + 1
            car$ = Mid$(chaine$, ich%, 1)
            If car$ = ")" Then
               compteur% = compteur% - 1
            ElseIf car$ = "(" Then
               compteur% = compteur% + 1
            End If
            If compteur% = 0 Then
               Exit Do
            End If
         Loop
      ElseIf car$ = "^" Then
         ' -------------------------------------
         ' On a une expression de la forme u^v :
         ' ----------------------------------
         iexpon% = ich%
         Exit Do
      End If
      ich% = ich% + 1
      If ich% > longchaine% Then
         Exit Do
      End If
   Loop
   ' ---------------------------------------------
   ' Examen de l'expression pour
   ' rechercher le type de u
   ' et renvoi de u' si expr. = u
   ' ou  (u^v)' = u^v*v'*LOG(u) + u^(v-1)*u'*v
   ' si expr. = u^v
   ' ---------------------------------------------
   ' ------------------
   ' Détermination de u
   ' ------------------
   FoFlag% = 0
   idebut% = 1
   If iexpon% = 0 Then
      ifin% = longchaine%
   Else
      ifin% = iexpon% - 1
   End If
   ufoncpar$ = Mid(chaine$, idebut%, ifin%)
   ' ----------------------------------
   ' Recherche d'une fonction numérique
   ' ----------------------------------
   ich% = 1
   troiscar$ = Mid$(ufoncpar$, ich%, 4)
   For iFoNum% = 1 To 11
      If troiscar$ = FoNum$(iFoNum%) Then
         FoFlag% = iFoNum%
         Exit For
      End If
   Next iFoNum%
   If FoFlag% <> 0 Then
      ' ----------------------------------
      ' on a trouvé une fonction numérique
      ' ----------------------------------
      argfonc$ = Mid(ufoncpar$, 5, ifin% - 5)
      Call DeriveSomme(argfonc$, varloc$, argfoncderive$)
      Select Case FoFlag%
      Case 1
         '  ufoncpar$ = "ABS(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "SGN(" & argfonc$ & ")"
         Else
            ufoncparderive$ = "SGN(" & argfonc$ & ")*(" & argfoncderive$ & ")"
         End If
      Case 2
         '  ufoncpar$ = "ATN(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "1/(1+(" & argfonc$ & ")^2)"
         Else
            ufoncparderive$ = "(" & argfoncderive$ & ")/(1+(" & argfonc$ & ")^2)"
         End If
      Case 3
         '  ufoncpar$ = "COS(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "-SIN(" & argfonc$ & ")"
         Else
            ufoncparderive$ = "-SIN(" & argfonc$ & ")*(" & argfoncderive$ & ")"
         End If
      Case 4
         '  ufoncpar$ = "EXP(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "EXP(" & argfonc$ & ")"
         Else
            ufoncparderive$ = "EXP(" & argfonc$ & ")*(" & argfoncderive$ & ")"
         End If
      Case 5
         '  ufoncpar$ = "FIX(" & argfonc$ & ")"
         Erreur = True
         Message$ = "La fonction FIX n'est pas dérivable !"
         MsgBox Message$, 48, "DERIVE"
         Exit Sub
      Case 6
         '  ufoncpar$ = "INT(" & argfonc$ & ")"
         Erreur = True
         Message$ = "La fonction INT n'est pas dérivable !"
         MsgBox Message$, 48, "DERIVE"
         Exit Sub
      Case 7
         '  ufoncpar$ = "LOG(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "1/(" & argfonc$ & ")"
         Else
            ufoncparderive$ = "(" & argfoncderive$ & ")/(" & argfonc$ & ")"
         End If
      Case 8
         '  ufoncpar$ = "SGN(" & argfonc$ & ")"
         ufoncparderive$ = "0"
      Case 9
         '  ufoncpar$ = "SIN(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "COS(" & argfonc$ & ")"
         Else
            ufoncparderive$ = "COS(" & argfonc$ & ")*(" & argfoncderive$ & ")"
         End If
      Case 10
         '  ufoncpar$ = "SQR(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "1/(2*SQR(" & argfonc$ & "))"
         Else
            ufoncparderive$ = "(" & argfoncderive$ & ")/(2*SQR(" & argfonc$ & "))"
         End If
      Case 11
         '  ufoncpar$ = "TAN(" & argfonc$ & ")"
         If argfoncderive$ = "1" Then
            ufoncparderive$ = "1/COS(" & argfonc$ & ")^2"
         Else
            ufoncparderive$ = "(" & argfoncderive$ & ")/COS(" & argfonc$ & ")^2"
         End If
      End Select
   Else
      ' ----------------------------------
      ' ça n'est pas une fonction;
      ' on cherche autre chose...
      ' ----------------------------------
      car$ = Mid$(ufoncpar$, ich%, 1)
      If car$ = "(" Then
         ' ----------------------------------
         ' on a trouvé une expression
         ' entre parenthèses
         ' ----------------------------------
         argpar$ = Mid(ufoncpar$, 2, ifin% - 2)
         Call DeriveSomme(argpar$, varloc$, argparderive$)
         On Error Resume Next
         nbep = CSng(argparderive$)
         If Err.Number = 0 Then
            ' v est un nombre
            ufoncparderive$ = argparderive$
         Else
            ' v est une constante ou contient la variable
            ufoncparderive$ = "(" & argparderive$ & ")"
         End If
         On Error GoTo 0
      ElseIf ufoncpar$ = varfonpar$ Then
         ' ----------------------------------
         ' on a trouvé la variable
         ' ----------------------------------
         ufoncparderive$ = "1"
      Else
         ' ----------------------------------
         ' la seule possibilité restante est
         ' que u soit une constante
         ' ----------------------------------
         ufoncparderive$ = "0"
      End If
   End If
   ' ---------------------------------------------
   ' Renvoi de u' si expr. = u
   ' ou  (u^v)' = u^v*v'*LOG(u) + u^(v-1)*u'*v
   ' si expr. = u^v
   ' ---------------------------------------------
   If iexpon% = 0 Then
      ' -----------------
      ' expr. = u
      ' -----------------
      exprfonparderivee$ = ufoncparderive$
   Else
      ' -----------------
      ' expr. = u^v
      ' -----------------
      vfoncpar$ = Mid(chaine$, iexpon% + 1, longchaine%)
      Call DeriveSomme(vfoncpar$, varloc$, vfoncparderive$)
      ' ------------------------------------
      ' examine les cas pouvant conduire
      ' à une simplification de l'écriture :
      ' (u^v)' = u^v*v'*LOG(u) + u^(v-1)*u'*v
      ' ------------------------------------
      If ufoncparderive$ = "0" And vfoncparderive$ = "0" Then
         ' u et v sont des nombres ou des constantes
         ' (u^v)' = 0
         exprfonparderivee$ = "0"
      ElseIf ufoncparderive$ = "0" Then
         ' u est un nombre ou une constante
         ' et v contient la variable
         If vfoncparderive$ = "1" Then
            ' v est la variable
            ' ( plus éventuellement une constante)
            ' (u^v)' = u^v*LOG(u)
            efpd$ = ufoncpar$ & "^" & vfoncpar$
            efpd$ = efpd$ & "*LOG(" & ufoncpar$ & ")"
            exprfonparderivee$ = efpd$
         Else
            ' (u^v)' = u^v*v'*LOG(u)
            efpd$ = ufoncpar$ & "^" & vfoncpar$ & "*" & vfoncparderive$
            efpd$ = efpd$ & "*LOG(" & ufoncpar$ & ")"
            exprfonparderivee$ = efpd$
         End If
      ElseIf vfoncparderive$ = "0" Then
         ' v est un nombre ou une constante
         ' et u contient la variable
         On Error Resume Next
         nbex = CSng(vfoncpar$)
         If nbex = 0 Or Err.Number <> 0 Then
            ' v est une constante
            efpd$ = ufoncpar$ & "^(" & vfoncpar$ & "-1)"
         Else
            ' v est un nombre
            If nbex - 1 = 1 Then
               efpd$ = ufoncpar$
            Else
               efpd$ = ufoncpar$ & "^" & Trim(Str(nbex - 1))
            End If
         End If
         On Error GoTo 0
         If ufoncparderive$ = "1" Then
            ' u est la variable
            ' ( plus éventuellement une constante)
            ' (u^v)' = u^(v-1)*v
            efpd$ = efpd$ & "*" & vfoncpar$
         Else
            ' (u^v)' = u^(v-1)*u'*v
            efpd$ = efpd$ & "*" & ufoncparderive$ & "*" & vfoncpar$
         End If
         exprfonparderivee$ = efpd$
      Else
         ' u et v contiennent la variable
         If ufoncparderive$ = "1" And vfoncparderive$ = "1" Then
            ' u et v sont la variable
            ' (u^v)' = u^v*LOG(u) + u^(v-1)*v
            efpd$ = ufoncpar$ & "^" & vfoncpar$
            efpd$ = efpd$ & "*LOG(" & ufoncpar$ & ")+"
            efpd$ = efpd$ & ufoncpar$ & "^(" & vfoncpar$
            efpd$ = efpd$ & "-1)*" & vfoncpar$
            exprfonparderivee$ = efpd$
         ElseIf ufoncparderive$ = "1" Then
            ' u est la variable
            ' ( plus éventuellement une constante)
            ' (u^v)' = u^v*v'*LOG(u) + u^(v-1)*v
            efpd$ = ufoncpar$ & "^" & vfoncpar$ & "*" & vfoncparderive$
            efpd$ = efpd$ & "*LOG(" & ufoncpar$ & ")+"
            efpd$ = efpd$ & ufoncpar$ & "^(" & vfoncpar$
            efpd$ = efpd$ & "-1)*" & vfoncpar$
            exprfonparderivee$ = efpd$
         ElseIf vfoncparderive$ = "1" Then
            ' v est la variable
            ' ( plus éventuellement une constante)
            ' (u^v)' = u^v*LOG(u) + u^(v-1)*u'*v
            efpd$ = ufoncpar$ & "^" & vfoncpar$
            efpd$ = efpd$ & "*LOG(" & ufoncpar$ & ")+"
            efpd$ = efpd$ & ufoncpar$ & "^(" & vfoncpar$
            efpd$ = efpd$ & "-1)*" & ufoncparderive$ & "*" & vfoncpar$
            exprfonparderivee$ = efpd$
         Else
            ' Cas général
            ' (u^v)' = u^v*v'*LOG(u) + u^(v-1)*u'*v
            efpd$ = ufoncpar$ & "^" & vfoncpar$ & "*" & vfoncparderive$
            efpd$ = efpd$ & "*LOG(" & ufoncpar$ & ")+"
            efpd$ = efpd$ & ufoncpar$ & "^(" & vfoncpar$
            efpd$ = efpd$ & "-1)*" & ufoncparderive$ & "*" & vfoncpar$
            exprfonparderivee$ = efpd$
         End If
      End If
   End If
End Sub

Public Sub InvMatCholeski(OrdreMsdp%, DefPos, DetMsdp, Msdp(), Wsdp())
' -----------------------------------------------------------------------
' Inversion d'une matrice carrée symétrique définie positive Msdp
' Méthode de Choleski
' ----------
' David BLUM
' 28/02/1997
' ----------
' En entrée :
' OrdreMsdp%   = ordre de la matrice Msdp
' Msdp()       = matrice à inverser
' En sortie :
' Wsdp()       = matrice inverse
' DetMsdp      = déterminant de Msdp
' DefPos       = True si la matrice est définie positive, donc inversible
'                False sinon
' -------------------------------------------------------------------
' On met la matrice sous la forme :
' Msdp = L x LTRANS
' où L est une matrice triangulaire inférieure et LTRANS sa transposée.
' On a :
' (Msdp)INV = (LTRANS)INV x LINV
' ---------------------------------------------------------------------
' Remarque :
' ----------
' La matrice étant symétrique, on n'utilise que
' ses éléments M(i,j) tels que j>i :
' M(1,1),M(1,2),M(1,3),...,M(1,n)
' M(2,2),M(2,3),...,M(2,n)
' ................................
' M(i,i),M(i,i+1),...,M(i,n)
' ................................
' M(n,n)
' ---------------------------------------------
' Initialisations
' ---------------
eps = 0.0001
Erreur = False
DetMsdp = 1
DefPos = True
' ---------------------------------------------------------
' Ordre de la matrice
' ---------------------------------------------------------
n% = OrdreMsdp%
If n% < 2 Then
   Beep
   MsgBox "L'ordre de la matrice doit être au moins égal à 2 !", 48, "InvMatCholeski"
   Exit Sub
End If
' ---------------------------------------------------------
' Dims et éléments de la matrice M
' -----------------------------------------------------------------
ReDim Wsdp(1 To n%, 1 To n%)      ' inverse de M
ReDim Lmat(1 To n%, 1 To n%)      ' matrice triangulaire inférieure L
ReDim LINVmat(1 To n%, 1 To n%)   ' inverse de L
' -----------------------------------------------------------------
' ***********
' Calcul de L
' ***********
' On fait le produit L x LTRANS puis on identifie
' le résultat à la matrice M
' -----------------------------------------------
For i% = 1 To n%
   For j% = i% To n%
      Lint = 0
      If i% > 1 Then
         For k% = 1 To i% - 1
            Lint = Lint + Lmat(i%, k%) * Lmat(j%, k%)
         Next k%
      End If
      If j% = i% Then
         li2 = Msdp(i%, i%) - Lint
         If li2 < eps Then
            DefPos = False
            Erreur = True
            Exit Sub
         End If
         Lmat(i%, i%) = Sqr(li2)
         DetMsdp = DetMsdp * Lmat(i%, i%)
      Else
         Lmat(i%, j%) = 0
         Lmat(j%, i%) = (Msdp(i%, j%) - Lint) / Lmat(i%, i%)
      End If
   Next j%
Next i%
' ----------------------------------------------------------------
' **************
' Calcul de LINV
' **************
' On fait le produit L x LINV puis on identifie
' le résultat à la matrice unité I
' ---------------------------------------------
' éléments diagonaux de I égaux à 1 :
For i% = 1 To n%
   LINVmat(i%, i%) = 1 / Lmat(i%, i%)
Next i%
' ----------------
' éléments non diagonaux de I égaux à 0 :
For j% = 1 To n% - 1
   For i% = j% + 1 To n%
      Lint = 0
      For k% = j% To i% - 1
         Lint = Lint + Lmat(i%, k%) * LINVmat(k%, j%)
      Next k%
      LINVmat(i%, j%) = -Lint / Lmat(i%, i%)
      LINVmat(j%, i%) = 0
   Next i%
Next j%
' ----------------
' ----------------------------------------------------------------
' ************************
' Calcul de W inverse de M
' ************************
' On a :
' (Msym)INV = (LTRANS)INV x LINV = (LINV)TRANS x LINV
' ----------------
For i% = 1 To n%
   For j% = i% To n%
      Mint = 0
      For k% = 1 To n%
         Mint = Mint + LINVmat(k%, i%) * LINVmat(k%, j%)
      Next k%
      Wsdp(i%, j%) = Mint
      Wsdp(j%, i%) = Mint
   Next j%
Next i%
' **************************
' Calcul du déterminant de M
' **************************
DetMsdp = DetMsdp * DetMsdp
End Sub

Public Sub VerifieMatSym(OrdreMsym%, Msym())
' -------------------------------------------------------------
' Vérifie que la matrice Msym est symétrique
' David BLUM
' 07/09/1997
' ----------------------------------------------------------------
' Initialisations
eps = 0.0001
Erreur = False
' ---------------------------------------------------------
' Ordre de la matrice
' ---------------------------------------------------------
n% = OrdreMsym%
If n% < 2 Then
   Beep
   Message$ = "L'ordre de la matrice doit être au moins égal à 2 !"
   MsgBox Message$, 48, "VerifieMatSym"
   Exit Sub
End If
' ---------------------------------------------------------
' Vérification
' -----------------------------------------------------------------
For i% = 1 To n%
   For j% = 1 To n%
      If i% <> j% Then
         difij = Abs(Msym(i%, j%) - Msym(j%, i%))
         If difij > eps Then
            Erreur = True
            Message$ = "La matrice n'est pas symétrique !"
            MsgBox Message$, 48, "VerifieMatSym"
            Exit Sub
         End If
      End If
   Next j%
Next i%
' ----------------------------------------------------------------
End Sub

Public Sub ZerosPolMatComp()
' -------------------------------------------------------------
' Recherche des racines d'un polynôme
' par la méthode de la matrice compagne
' David BLUM
' 27/09/1997
' -------------------------------------------------------------
' Recherche les racines du polynôme Ppol de degré DegPpol%
' S'il y a erreur ou non convergence, fait Erreur%=1 et sort;
' Sinon remplit ModuleRac(), RacineR() et RacineI()
' avec les modules et les parties réelles
' et imaginaires des racines triées par ordre de
' module décroissant.
' ----------------------------------------------------------------
' Polynôme Ppol(x) d'ordre n :
' Ppol(x) = a(n) * x^n + a(n-1) * x^(n-1) + ... + a(1) * x + a(0)
' ----------------------------------------------------------------
' Matrice compagne Mcomp : Mcomp(1,j)   = -a(n-j)/a(n)
'                          Mcomp(i+1,i) = 1
'               les autres Mcomp(i,j) = 0
' ----------------------------------------------------------------------
' Ppol(x) est, au facteur (-1)^n près, le polynôme caractéristique
' de Mcomp. On recherche les valeurs propres (réelles et complexes)
' de Mcomp par la méthode du double QR; ce sont aussi les zéros de Ppol.
' ----------------------------------------------------------------------
' Initialisation
' ---------------------------------------------------------
Erreur = False
eps = 0.000001
' ---------------------------------------------------------
' Degré suffisant ?
' ---------------------------------------------------------
If DegPpol% < 1 Then
   Erreur = True
   Message$ = " Polynôme de degré inférieur à 1 ! "
   MsgBox Message$, 16, "ZerosPolMatComp"
   Exit Sub
End If
' ---------------------------------------------------------
' Ppol(n) <> 0 ?
' ---------------------------------------------------------
If Abs(Ppol(DegPpol%)) < eps Then
   Erreur = True
   Message$ = " Polynôme de degré inférieur à " & DegPpol% & " ! "
   MsgBox Message$, 16, "ZerosPolMatComp"
   Exit Sub
End If
' ---------------------------------------------------------
' Dims
' ---------------------------------------------------------
ReDim Mcomp(1 To DegPpol%, 1 To DegPpol%)
ReDim Gcomp(1 To DegPpol%, 1 To DegPpol%)
ReDim PermComp%(1 To DegPpol%)
ReDim RacineR(1 To DegPpol%)
ReDim RacineI(1 To DegPpol%)
ReDim ModuleRac(1 To DegPpol%)
' ---------------------------------------------------------
' Initialisations
' ---------------------------------------------------------
For iloc% = 1 To DegPpol%
   RacineR(iloc%) = 0
   RacineI(iloc%) = 0
   ModuleRac(iloc%) = 0
   For jloc% = 1 To DegPpol%
      Mcomp(iloc%, jloc%) = 0
   Next jloc%
Next iloc%
' ---------------------------------------------------------
' Construction de Mcomp
' ---------------------------------------------------------
' Matrice compagne Mcomp : Mcomp(1,j)   = -a(n-j)/a(n)
'                          Mcomp(i+1,i) = 1
'               les autres Mcomp(i,j) = 0
' ---------------------------------------------------------
For iloc% = 1 To DegPpol%
   Mcomp(1, iloc%) = -Ppol(DegPpol% - iloc%) / Ppol(DegPpol%)
Next iloc%
For iloc% = 1 To DegPpol% - 1
   Mcomp(iloc% + 1, iloc%) = 1
Next iloc%
' -----------------------------------
' Réduction de M sous forme
' d'une matrice de HESSENBERG G
' -----------------------------------
Call FenetreMatrice.Hessenberg(DegPpol%, Mcomp(), Gcomp(), PermComp%())
Erase PermComp%
' ---------------------------------------------------------
' Appel de DoubleQR
' ---------------------------------------------------------
Call FenetreMatrice.DiaMaQR2Vap(DegPpol%, NbZerosComp%, Gcomp(), RacineR(), RacineI())
Erase Mcomp, Gcomp
' -----------------------------------------------------------------
' Calcul des modules des zéros de Ppol(x) :
' -----------------------------------------------------------------
For iloc% = 1 To DegPpol%
   ModuleRac(iloc%) = Sqr(RacineR(iloc%) ^ 2 + RacineI(iloc%) ^ 2)
Next iloc%
' -----------------------------------------------------------------
' Classement des zéros de Ppol(x) selon leur modules décroissants :
' -----------------------------------------------------------------
For numzer% = 1 To DegPpol% - 1
   ModuleRac0 = ModuleRac(numzer%)
   For iloc% = numzer% + 1 To DegPpol%
      ModuleRac1 = ModuleRac(iloc%)
      If ModuleRac0 < ModuleRac1 Then
         ' Echange de RacineR(numzer%) et RacineR(iloc%)
         RacineR1 = RacineR(iloc%)
         RacineR(iloc%) = RacineR(numzer%)
         RacineR(numzer%) = RacineR1
         ' Echange de RacineI(numzer%) et RacineI(iloc%)
         RacineI1 = RacineI(iloc%)
         RacineI(iloc%) = RacineI(numzer%)
         RacineI(numzer%) = RacineI1
         ' Echange de ModuleRac(numzer%) et ModuleRac(iloc%)
         ModuleRac2 = ModuleRac1
         ModuleRac1 = ModuleRac0
         ModuleRac0 = ModuleRac2
         ModuleRac(numzer%) = ModuleRac0
         ModuleRac(iloc%) = ModuleRac1
      End If
    Next iloc%
Next numzer%
' ---------------------------------------------------------
End Sub

Public Sub TriParInsertion(NbDeNb%, NbEnDesordre(), NbEnOrdre(), PermTpi%())
' **************************
' *    Tri par insertion   *
' **************************
' --------------------------------------------------------------------------
' Principe : On prend le premier nombre; on prend le deuxième nombre et on le
' classe en fonction du premier; on prend le troisième nombre et on le classe
' en en fonction des deux premiers; et ainsi de suite.
' Nombre de comparaisons à effectuer pour n nombres : environ n(n-1)/2
' --------------------------------------------------------------------------
' En entrée :     NbDeNb%        =  Nombre de nombres à trier
'                 NbEnDesordre() =  Liste des nombres à trier
'
' En sortie :     NbEnOrdre()    =  Liste des nombres triés
'                 PermTpi%()     =  Permutations effectuées lors du tri
'                                   [PermTpi%(j%) = ancienne position du nombre
'                                    placé après tri en position j%]
' --------------------------------------------------------------------------
If NbDeNb% < 2 Then
   Message$ = " Le nombre de nombres à trier doit être supérieur à 1 !"
   MsgBox Message$, 48, "TriParInsertion"
   Exit Sub
End If
' ---------------------------------------------------
' ********************************
' Algorithme de tri
' ********************************
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
For i% = 2 To NbDeNb%
   NbEnOrdre(i%) = 0
   PermTpi%(i%) = i%
Next i%
' ---------------------------------------------------
NbEnOrdre(1) = NbEnDesordre(1)
PermTpi%(1) = 1
For i% = 2 To NbDeNb%
   NbIns = NbEnDesordre(i%)
   NbEnOrdre(i%) = NbIns
   For j% = 1 To i% - 1
      If NbIns < NbEnOrdre(j%) Then
         For k% = i% - 1 To j% Step -1
            NbEnOrdre(k% + 1) = NbEnOrdre(k%)
            PermTpi%(k% + 1) = PermTpi%(k%)
         Next k%
         NbEnOrdre(j%) = NbIns
         PermTpi%(j%) = i%
         Exit For
      End If
   Next j%
Next i%
' --------------------------------------------------
End Sub

Public Sub TriRapide(NbDeNb%, NbEnDesordre(), NbEnOrdre(), PermTra%())
' *******************
' *    Tri-rapide   *
' *******************
' --------------------------------------------------------------------------
' Principe : On prend le premier nombre (appelé pivot); on compare les autres
' nombres au pivot et on les classe en 2 groupes : le premier constitué des
' nombres qui lui sont inférieurs et le second de ceux qui lui sont supérieurs;
' ce tri partiel effectué,on prend alors le premier groupe et on renouvelle
' avec lui la manipulation précedente, de même avec le deuxième groupe, etc...
' Nombre de comparaisons à effectuer pour n nombres : environ n.(log à base 2)(n)
'
' C'est la borne la meilleure que l'on puisse atteindre en n'utilisant que des
' comparaisons.
' --------------------------------------------------------------------------
' En entrée :     NbDeNb%        =  Nombre de nombres à trier
'                 NbEnDesordre() =  Liste des nombres à trier
'
' En sortie :     NbEnOrdre()    =  Liste des nombres triés
'                 PermTra%()     =  Permutations effectuées lors du tri
'                                   [PermTpi%(j%) = ancienne position du nombre
'                                    placé après tri en position j%]
' --------------------------------------------------------------------------
If NbDeNb% < 2 Then
   Message$ = " Le nombre de nombres à trier doit être supérieur à 1 !"
   MsgBox Message$, 48, "TriRapide"
   Exit Sub
End If
' ---------------------------------------------------
' Initialisations
' ---------------------------------------------------
For i% = 2 To NbDeNb%
   PermTra%(i%) = i%
Next i%
' ---------------------------------------------------
' ********************************
' Algorithme de tri
' ********************************
Call TriPartiel(1, NbDeNb%, 1, NbEnOrdre(), PermTra%())
' --------------------------------------------------
End Sub

Public Sub TriPartiel(debpa%, finpa%, pospivpa%, Nbpa(), PermTpa%())
' ------------------------------------------------------
' Tri partiel de nbnb% nombres
' ----------------------------------------- ------------
' En entrée :  debpa%      = début de la zone à trier
'              finpa%      = fin de la zone à trier
'              Nbpa()      = tableau des nombres à trier
'              PermTra%()  =  Permutations effectuées lors du tri
'                             [PermTpi%(j%) = ancienne position du nombre
'                              placé après tri en position j%]
'
' En sortie :  pospivpa%   = position du pivot
'              Nbpa()      = tableau des nombres triés
'              PermTra%()
' ------------------------------------------------------
NbNb% = finpa% - debpa% + 1
If NbNb% < 2 Then Exit Sub
ReDim Nbloc(1 To NbNb%)
ReDim PermLoc%(1 To NbNb%)
' ------------------------------------------------------
' Algorithme de tri partiel
' ------------------------------------------------------
pivotpa = Nbpa(debpa%)
ipivpa% = 1
Nbloc(ipivpa%) = pivotpa
PermLoc%(ipivpa%) = 1
For iloc% = 2 To NbNb%
   inbpa% = debpa% + iloc% - 1
   If Nbpa(inbpa%) >= pivotpa Then
      Nbloc(iloc%) = Nbpa(inbpa%)
      PermLoc%(iloc%) = iloc%
   Else
      For jloc% = iloc% To ipivpa% + 1 Step -1
         Nbloc(jloc%) = Nbloc(jloc% - 1)
         PermLoc%(jloc%) = PermLoc%(jloc% - 1)
      Next jloc%
      Nbloc(ipivpa%) = Nbpa(inbpa%)
      PermLoc%(ipivpa%) = iloc%
      ipivpa% = ipivpa% + 1
   End If
Next iloc%
pospivpa% = debpa% + ipivpa% - 1
' ------------------------------------------------------
' Transfert de Nbloc à Nbpa et PermLoc% à PermTpa%
' ------------------------------------------------------
For iloc% = 1 To NbNb%
   Nbpa(debpa% + iloc% - 1) = Nbloc(iloc%)
   PermTpa%(debpa% + iloc% - 1) = PermLoc%(iloc%)
Next iloc%
' ----------------------------------------------------------
' Libération de la mémoire occupée par Nbloc() et PermLoc%()
' ----------------------------------------------------------
Erase Nbloc
Erase PermLoc%
' ------------------------------------------------------
' Tri partiel des 2 sous-ensembles formés
' ------------------------------------------------------
Call TriPartiel(debpa%, pospivpa% - 1, pospivpa1%, Nbpa(), PermTpa%())
Call TriPartiel(pospivpa% + 1, finpa%, pospivpa2%, Nbpa(), PermTpa%())
' ------------------------------------------------------
End Sub
