Imports System.IO
Public Class WinQuinte
    Dim Chevaux As ArrayList
    Dim ChxTri As ArrayList
    Dim ChxSyn As ArrayList
    Dim ChxElim As ArrayList
    Dim Quintes As ArrayList
    Dim TBX(4) As Integer
    Dim T(4, 16) As Integer
    Dim B(13) As Integer
    Dim B5(4, 13) As Integer
    Dim B3(2, 13) As Integer
    Dim B9(8, 1500) As Integer
    Dim BX(4, 999) As Integer
    Dim BXT(4, 999) As Integer
    Dim BXS(4, 999) As Integer
    Dim TRI(4, 999) As Integer
    Dim ARQT(4) As Integer
    Dim EL(13) As Integer
    Dim BQ(2) As Integer
    Dim CH(13) As Integer
    Dim QT(999) As String
    Dim QTN(4) As Integer
    Dim SX(6) As Integer
    Dim SX2(4, 5) As Integer
    Dim SX3(4, 20) As Integer
    Dim NF(8) As Integer
    Dim NF1(8) As Integer
    Dim NF2(8) As Integer
    Dim T1(4) As Integer
    Dim P1(4) As Integer
    Dim FLGELI As Integer
    Dim ELI As Integer
    Dim RES As Integer
    Dim ELIS As Integer
    Dim RESS As Integer
    Dim BAS As Integer
    Dim DateCourse As String
    Dim DatePM As Integer
    Dim NumFile As Integer
    Dim MaxFile As Integer
    Dim CPTC As Integer
    Dim NB As Integer
    Dim TPage As Integer
    Dim CPage As Integer
    Dim IBox As Integer
    Dim CBox As Integer
    Dim LBox As Integer
    Dim ord As Integer
    Dim EG As Integer
    Dim MyFile As StreamReader
    Dim TRIFile As StreamReader
    Dim Ligne(21) As String
    Dim LQ(4) As String
    Dim MyFileW As StreamWriter
    Dim TXTFileW As StreamWriter
    Dim TRIFileW As StreamWriter
    Dim DateCourante As Date
    Dim Pos As Integer
    Dim SCH As Integer
    Dim Ok As Boolean
    Dim SELECFlag As Boolean
    '***
    '*** CHARGEMENT
    '***

    Private Sub WinQuinte_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '*** activation du logiciel WINQUINT
        PictureBox1.BackgroundImage = Image.FromFile("GranvillaiseBleue.bmp")
        PictureBox2.BackgroundImage = Image.FromFile("GranvillaiseBleue.bmp")
        PictureBox3.BackgroundImage = Image.FromFile("GranvillaiseBleue.bmp")
        PictureBox4.BackgroundImage = Image.FromFile("GranvillaiseBleue.bmp")
        '*** 
        DatePM = 0
        DateCourante = Now
        DateCourse = Format(DateAdd("d", DatePM, DateCourante), "yyyyMMdd")
        DateJour.Text = Format(DateAdd("d", DatePM, DateCourante), "Long Date")
        ANNUL.Enabled = False
        '*** 
        SELECFlag = False
        '***
        PAGEAide.Navigate("")
        '***
        Call ChargementREF()
    End Sub

    '***
    '*** SYNTHESE
    '***

    Private Sub DateMoins_Click(sender As Object, e As EventArgs) Handles DateMoins.Click
        DatePM = -1
        Call DateJour_Changed()
    End Sub

    Private Sub DatePlus_Click(sender As Object, e As EventArgs) Handles DatePlus.Click
        DatePM = 1
        Call DateJour_Changed()
    End Sub

    Private Sub DateJour_Changed()
        DateCourse = Format(DateAdd("d", DatePM, DateCourante), "yyyyMMdd")
        DateJour.Text = Format(DateAdd("d", DatePM, DateCourante), "Long Date")
        DateCourante = DateAdd("d", DatePM, DateCourante)
        ANNUL.Enabled = False
        '*** 
        SELECFlag = False
        '***
        Call ChargementREF()
    End Sub

    Sub ChargementREF()
        Dim I As Integer
        '*** 
        Call Initialisation()
        '*** 
        '*** Chargement des références de la course correspondants à la date du jour
        '*** 
        If File.Exists(DateCourse & ".REF") Then
            WINQUINT.TabPages.Item("ELIMINATION").Enabled = True
            WINQUINT.TabPages.Item("VISUALISATION").Enabled = True
            MyFile = My.Computer.FileSystem.OpenTextFileReader(DateCourse & ".REF")
            Ligne = Split(MyFile.ReadLine(), ",")
            For I = 0 To 21
                If I < 14 Then
                    B(I) = Ligne(I)
                    Chevaux(I).text = Ligne(I)
                ElseIf I < 17 Then
                    BQ(I - 14) = Ligne(I)
                Else
                    ARQT(I - 17) = Ligne(I)
                End If
            Next I
            MyFile.Close()
            '***
            BA1.Text = BQ(0)
            BA2.Text = BQ(1)
            BA3.Text = BQ(2)
            '*** 
            If BQ(2) > 0 Then
                NbBase3.Checked = True
                NB = 2
            Else
                If BQ(1) > 0 Then
                    NbBase2.Checked = True
                    BA3.Visible = False
                    NB = 1
                Else
                    NbBase1.Checked = True
                    BA2.Visible = False
                    BA3.Visible = False
                    NB = 0
                End If
            End If
            If ARQT(0) > 0 Then
                QTA1.Text = ARQT(0)
                QTA2.Text = ARQT(1)
                QTA3.Text = ARQT(2)
                QTA4.Text = ARQT(3)
                QTA5.Text = ARQT(4)
            End If
            '***
            NumFile = 0 '***Fichier AAAAMMJJ.TR0
            Call ChargementListeQuintes()
        Else
            SELMOINS.Enabled = False
            SELPLUS.Enabled = False
            DateJour.Focus()
            If Not IsNothing(WINQUINT.TabPages.Item("ELIMINATION")) Then
                WINQUINT.TabPages.Item("ELIMINATION").Enabled = False
            End If
            If Not IsNothing(WINQUINT.TabPages.Item("VISUALISATION")) Then
                WINQUINT.TabPages.Item("VISUALISATION").Enabled = False
            End If
        End If
        '***
    End Sub

    Private Sub SELMOINS_Click(sender As Object, e As EventArgs) Handles SELMOINS.Click
        NumFile -= 1
        If NumFile <= 0 Then
            NumFile = 0
        End If
        Call ChargementListeQuintes()
    End Sub

    Private Sub SELPLUS_Click(sender As Object, e As EventArgs) Handles SELPLUS.Click
        NumFile += 1
        If NumFile >= 9 Then
            COPY.Enabled = False
            NumFile = 9
        End If
        Call ChargementListeQuintes()
    End Sub

    Private Sub SELEC_TextChanged(sender As Object, e As EventArgs) Handles SELEC.TextChanged
        Dim I As Integer
        If SELECFlag = False Then Exit Sub
        If Not IsNothing(DateCourse) Then
            For I = 0 To 9
                If Not File.Exists(DateCourse & ".TR" & I) Then
                    Exit For
                End If
            Next I
            Select Case I
                Case 1
                    SELMOINS.Enabled = False
                    SELPLUS.Enabled = False
                Case Is > 1
                    If NumFile = I - 1 Then
                        SELMOINS.Enabled = True
                        SELPLUS.Enabled = False
                    Else
                        SELMOINS.Enabled = True
                        SELPLUS.Enabled = True
                    End If
            End Select
        End If
    End Sub

    Private Sub NbBase1_CheckedChanged(sender As Object, e As EventArgs) Handles NbBase1.CheckedChanged
        NB = 0
        BA1.Visible = True
        If BQ(0) < 1 Then
            BA1.Text = ""
            BA1.Focus()
        End If
        BA2.Visible = False
        BA3.Visible = False
    End Sub

    Private Sub NbBase2_CheckedChanged(sender As Object, e As EventArgs) Handles NbBase2.CheckedChanged
        NB = 1
        BA2.Visible = True
        If BQ(0) < 1 Then
            BA2.Text = ""
            BA1.Focus()
        Else
            If BQ(1) < 1 Then
                BA2.Text = ""
                BA2.Focus()
            End If
        End If
        BA3.Visible = False
    End Sub

    Private Sub NbBase3_CheckedChanged(sender As Object, e As EventArgs) Handles NbBase3.CheckedChanged
        NB = 2
        BA2.Visible = True
        BA3.Visible = True
        If BQ(0) < 1 Then
            BA2.Text = ""
            BA3.Text = ""
            BA1.Focus()
        Else
            If BQ(1) < 1 Then
                BA2.Text = ""
                BA3.Text = ""
                BA2.Focus()
            Else
                If BQ(2) < 1 Then
                    BA3.Text = ""
                    BA3.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub Generation_Click(sender As Object, e As EventArgs) Handles Generation.Click
        Dim I As Integer
        Dim J As Integer
        Dim X As Integer
        Dim K As Integer
        Dim V As Integer
        Dim TX As Integer
        Dim PX As Integer
        Dim N As Integer
        Dim TJ As Integer
        Dim FI As Integer
        Dim IM As Integer
        Dim RI As Integer
        Dim flg As Integer
        Dim flgp As Integer
        Dim flgi As Integer
        Dim flgn As Integer
        Dim flg3 As Integer
        Dim flgE As Integer
        Dim CPT As Integer
        Dim cpt_31 As Integer
        Dim cpt_32 As Integer
        Dim cpt_33 As Integer
        Dim cpt_l9_31 As Integer
        Dim cpt_l9_32 As Integer
        Dim cpt_l9_ok As Integer
        Dim X5 As Integer
        Dim Y5 As Integer
        Dim X3 As Integer
        Dim Y3 As Integer
        Dim SY1 As Integer
        Dim SY2 As Integer
        Dim Y9 As Integer
        Dim TI As Integer
        Dim QTM As String
        Dim REPONSE As Integer
        '*** Gestion des bases
        If NbBase1.Checked = True Then
            BQ(1) = 0
            BQ(2) = 0
        Else
            If NbBase2.Checked = True Then
                BQ(2) = 0
            End If
        End If
        '*** Test de validite des numéros de chevaux saisis
        For I = 0 To 13
            Ok = True
            SCH = Val(Chevaux(I).Text)
            Pos = I + 1
            Call ChevauxVerifications()
            If Ok = False Then
                Chevaux(I).SelectAll()
                Chevaux(I).Focus()
                Exit Sub
            Else
                B(I) = Val(Chevaux(I).Text)
            End If
        Next I
        '*** Test de validite des numéros de base(s) saisie(s)
        SCH = Val(BA1.Text)
        Ok = True
        Call BasesVerifications()
        If Ok = False Then
            BA1.SelectAll()
            BA1.Focus()
            Exit Sub
        Else
            BQ(0) = Val(BA1.Text)
        End If
        If NbBase1.Checked = True Then
            GoTo REGEN
        End If
        SCH = Val(BA2.Text)
        Ok = True
        Call BasesVerifications()
        If Ok = False Then
            BA2.SelectAll()
            BA2.Focus()
            Exit Sub
        Else
            BQ(1) = Val(BA2.Text)
        End If
        If NbBase2.Checked = True Then
            GoTo REGEN
        End If
        SCH = Val(BA3.Text)
        Ok = True
        Call BasesVerifications()
        If Ok = False Then
            BA3.SelectAll()
            BA3.Focus()
            Exit Sub
        Else
            BQ(2) = Val(BA3.Text)
        End If
REGEN:
        '*** Test de regeneration de base
        If SELMOINS.Enabled = True Or SELPLUS.Enabled = True Then
            REPONSE = MsgBox("Voulez-vous regénérer les combinaisons ?", 1, "REGENERATION DES COMBINAISONS")
            If REPONSE = 1 Then
                For I = 0 To 9
                    On Error Resume Next
                    Kill(DateCourse & ".TR" & I)
                Next I
                For I = 0 To 4
                    ARQT(I) = 0
                Next I
                QTA1.Text = ""
                QTA2.Text = ""
                QTA3.Text = ""
                QTA4.Text = ""
                QTA5.Text = ""
                GoTo VALID
            Else
                Exit Sub
            End If
        End If
VALID:
        '*** initialisation des tables
        For I = 0 To 1500
            For J = 0 To 8
                B9(J, I) = 0
            Next J
        Next I
        For I = 0 To 999
            For J = 0 To 4
                BXT(J, I) = 0
            Next J
        Next I
        Call ChargementGrilleSynthese()
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 1084
        ProgressBar1.Visible = True
        ProgressBar1.Value = ProgressBar1.Minimum
        '*** Génération des combinaisons
        '*** I/ constitution des base 9
        '***    1 ligne base 5 + 3 ligne base 3 <= 9 chevaux
        Y5 = 0
        '*** tant que les 14 ligne base 5 ne sont pas testées
        While Y5 < 14
            ProgressBar1.Value = Y9
            '*** chargement de la base 5 dans le debut
            '*** de la base 9 temporaire
            For I = 0 To 4
                NF(I) = B5(I, Y5)
                NF1(I) = B5(I, Y5)
                NF2(I) = B5(I, Y5)
            Next I
            For I = 5 To 8
                NF(I) = 0
                NF1(I) = 0
                NF2(I) = 0
            Next I
            '*** initialisation des indices et variables
            cpt_l9_ok = 4   'compteur de chevaux ok en base 9
            cpt_31 = 1      'compteur de ligne base 3.1
            SY1 = 0         'sauvegarde ligne 3.1
            SY2 = 0         'sauvegarde ligne 3.2
            X3 = 0
            Y3 = 0
            X5 = 0
            '*** tant que ligne base 3.1 pas toutes testees
            While Y3 < 12
                '*** tant que les 3 chevaux ligne
                '*** base 3.1 ne sont pas tous testés
                X3 = 0
                While X3 < 3
                    '*** test de validite du cheval ligne base 3
                    flg3 = 0
                    For I = 0 To 4
                        If NF(I) = B3(X3, Y3) Then
                            flg3 = 1
                            I = 4
                        End If
                    Next I
                    '*** si le cheval ligne 3 est valide
                    If flg3 = 0 Then
                        cpt_l9_ok += 1
                        NF(cpt_l9_ok) = B3(X3, Y3)
                    End If
                    '*** On passe au cheval suivant
                    '*** de la ligne 3.1
                    X3 += 1
                End While
                For I = 5 To 8
                    NF1(I) = NF(I)
                Next I
                cpt_l9_31 = cpt_l9_ok
                Y3 += 1
                SY1 = Y3
                '*** tant que ligne base 3.2 pas toutes testees
                While Y3 < 13
                    '*** tant que les 3 chevaux ligne
                    '*** base 3 ne sont pas tous testés
                    X3 = 0
                    While X3 < 3  '4
                        '*** test de validite du
                        '*** cheval ligne base 3
                        flg3 = 0
                        For I = 0 To 8
                            If NF(I) = B3(X3, Y3) Then
                                flg3 = 1
                                I = 8
                            End If
                        Next I
                        '*** si le cheval ligne 3 est valide
                        If flg3 = 0 Then
                            If cpt_l9_ok < 8 Then
                                cpt_l9_ok += 1
                                NF(cpt_l9_ok) = B3(X3, Y3)
                            End If
                        End If
                        '*** On passe au cheval suivant
                        '*** de la ligne 3
                        X3 += 1
                    End While
                    For I = 5 To 8
                        NF2(I) = NF(I)
                    Next I
                    cpt_l9_32 = cpt_l9_ok
                    Y3 += 1
                    SY2 = Y3
                    '*** tant que ligne base 3.3
                    '*** pas toutes testees
                    While Y3 < 14
                        '*** tant que les 3 chevaux ligne
                        '*** base 3 ne sont pas tous testés
                        X3 = 0
                        While X3 < 3
                            '*** test de validite du
                            '*** cheval ligne base 3
                            flg3 = 0
                            For I = 0 To 8
                                If NF(I) = B3(X3, Y3) Then
                                    flg3 = 1
                                    I = 8
                                End If
                            Next I
                            '*** si le cheval ligne 3
                            '*** est valide
                            If flg3 = 0 Then
                                If cpt_l9_ok < 8 Then
                                    cpt_l9_ok += 1
                                    NF(cpt_l9_ok) = B3(X3, Y3)
                                Else
                                    X3 = 2
                                End If
                            End If
                            '*** On passe au cheval suivant
                            '*** de la ligne 3
                            X3 += 1
                        End While
                        If cpt_l9_ok < 9 Then
                            '*** tant que les 9 chevaux
                            '*** ne sont pas dans l'ordre
                            CPT = 0
                            While CPT < 9
                                For I = 0 To 8
                                    For J = (I + 1) To 8
                                        If NF(I) <= NF(J) Then
                                            X = NF(J)
                                            NF(J) = NF(I)
                                            NF(I) = X
                                        End If
                                    Next J
                                Next I
                                CPT += 1
                            End While
                            '*** tant que toute la table base 9
                            '*** pas teste ou que base 9 tmp
                            '*** déja enregistree
                            I = 0
                            flg = 0
                            While I < 1500 And flg = 0
                                If B9(1, I) < 1 Or B9(1, I) > 20 Then
                                    flg = 2
                                Else
                                    CPT = 0
                                    For J = 0 To 8
                                        If B9(J, I) = NF(J) Then
                                            CPT += 1
                                        End If
                                    Next J
                                    If CPT = 9 Then
                                        flg = 1
                                    End If
                                    I += 1
                                End If
                            End While
                            If flg = 0 Or flg = 2 Then
                                For I = 0 To 8
                                    B9(I, Y9) = NF(I)
                                Next I
                                Y9 += 1
                                ProgressBar1.Value = Y9
                            End If
                        End If
                        Y3 += 1
                        For I = 0 To 8
                            NF(I) = NF2(I)
                        Next I
                        cpt_l9_ok = cpt_l9_32
                    End While
                    Y3 = SY2
                    For I = 0 To 8
                        NF(I) = NF1(I)
                        NF2(I) = NF1(I)
                    Next I
                    cpt_l9_ok = cpt_l9_31
                End While
                Y3 = SY1
                For I = 5 To 8
                    NF(I) = 0
                    NF1(I) = 0
                    NF2(I) = 0
                Next I
                cpt_l9_ok = 4
            End While
            Y5 += 1
        End While
        '*** Passage des bases 9 dans la BASES des 14 chevaux :
        I = 0
        TJ = 0
        ProgressBar2.Minimum = 0
        If NbBase3.Checked = True Then
            ProgressBar2.Maximum = 990
        Else
            If NbBase2.Checked = True Then
                ProgressBar2.Maximum = 750
            Else
                ProgressBar2.Maximum = 400
            End If
        End If
        ProgressBar2.Visible = True
        ProgressBar2.Value = ProgressBar2.Minimum
        '*** Tant que les lignes base 9 pas toutes testées
        While I < 999 And B9(0, I) > 0
            ProgressBar2.Value = TJ
            J = 0
            K = 0
            For CPT = 0 To 6
                SX(CPT) = 0
            Next CPT
            CPT = 0
            '*** Teste des BASES des 14 chevaux
            For J = 0 To 13
                flg = 0
                For K = 0 To 8
                    If B9(K, I) = B(J) Then
                        flg = 1
                    End If
                Next K
                If flg = 0 Then
                    SX(CPT) = B(J)
                    CPT += 1
                End If
            Next J
            CPT -= 1
            '*** Chargement des bases quinte si il reste 5 chevaux
            If CPT = 4 Then
                '*** Mise en ordre des chevaux
                For X5 = 0 To 4
                    For K = (X5 + 1) To 4
                        If SX(X5) <= SX(K) Then
                            X = SX(K)
                            SX(K) = SX(X5)
                            SX(X5) = X
                        End If
                    Next K
                Next X5
                '*** Elimination des quintes pairs, impairs
                '*** chevaux tous < 9 ou tous > 9
                flgp = 0 '*** Pairs
                flgi = 0 '*** Impairs
                flgn = 0 '*** Eliminés
                For X5 = 0 To 4
                    If SX(X5) = 2 Or SX(X5) = 4 Or SX(X5) = 6 Or SX(X5) = 8 Or SX(X5) = 10 Or
                        SX(X5) = 12 Or SX(X5) = 14 Or SX(X5) = 16 Or SX(X5) = 18 Or SX(X5) = 20 Then
                        flgp += 1
                    Else
                        flgi += 1
                    End If
                Next X5
                If flgp = 5 Or flgi = 5 Then
                    flgn = 1
                End If
                If SX(0) <= 9 And SX(1) <= 9 And SX(2) <= 9 And SX(3) <= 9 And SX(4) <= 9 Then
                    flgn = 1
                End If
                If SX(0) >= 10 And SX(1) >= 10 And SX(2) >= 10 And SX(3) >= 10 And SX(4) >= 10 Then
                    flgn = 1
                End If
                '*** tant que tous les quintes pas teste
                '*** ou que quinte déja enregistree
                If flgn = 0 Then
                    flg = 0
                    X5 = 0
                    While X5 < TJ And flg = 0
                        If BX(1, X5) = 0 Then
                            flg = 2
                        Else
                            CPT = 0
                            For K = 0 To 4
                                If BX(K, X5) = SX(K) Then
                                    CPT += 1
                                End If
                            Next K
                            If CPT = 5 Then
                                flg = 1
                            End If
                            X5 += 1
                        End If
                    End While
                    If flg = 0 Or flg = 2 Then
                        For X5 = 0 To 4
                            For V = 0 To NB
                                If SX(X5) = BQ(V) Then
                                    flg = 1
                                End If
                            Next V
                        Next X5
                        If flg = 1 Then
                            For X5 = 0 To 4
                                BX(X5, TJ) = SX(X5)
                            Next X5
                            TJ += 1
                        End If
                    End If
                End If
            Else
                '*** Chargement des bases quinte si il reste 6 chevaux
                If CPT = 5 Then
                    SX2(0, 0) = SX(0)
                    SX2(1, 0) = SX(1)
                    SX2(2, 0) = SX(2)
                    SX2(3, 0) = SX(3)
                    SX2(4, 0) = SX(4)
                    SX2(0, 1) = SX(0)
                    SX2(1, 1) = SX(1)
                    SX2(2, 1) = SX(2)
                    SX2(3, 1) = SX(3)
                    SX2(4, 1) = SX(5)
                    SX2(0, 2) = SX(0)
                    SX2(1, 2) = SX(1)
                    SX2(2, 2) = SX(2)
                    SX2(3, 2) = SX(4)
                    SX2(4, 2) = SX(5)
                    SX2(0, 3) = SX(0)
                    SX2(1, 3) = SX(1)
                    SX2(2, 3) = SX(3)
                    SX2(3, 3) = SX(4)
                    SX2(4, 3) = SX(5)
                    SX2(0, 4) = SX(0)
                    SX2(1, 4) = SX(2)
                    SX2(2, 4) = SX(3)
                    SX2(3, 4) = SX(4)
                    SX2(4, 4) = SX(5)
                    SX2(0, 5) = SX(1)
                    SX2(1, 5) = SX(2)
                    SX2(2, 5) = SX(3)
                    SX2(3, 5) = SX(4)
                    SX2(4, 5) = SX(5)
                    '*** tant que les 5 chevaux ne sont pas dans l'ordre
                    For Y5 = 0 To 5
                        For X5 = 0 To 4
                            For K = (X5 + 1) To 4
                                If SX2(X5, Y5) <= SX2(K, Y5) Then
                                    X = SX2(K, Y5)
                                    SX2(K, Y5) = SX2(X5, Y5)
                                    SX2(X5, Y5) = X
                                End If
                            Next K
                        Next X5
                    Next Y5
                    '*** tant que tous les quintes pas teste
                    '*** ou que quinte déja enregistree
                    Y5 = 0
                    While Y5 < 6
                        flg = 0
                        X5 = 0
                        While X5 < TJ And flg = 0
                            If BX(1, X5) = 0 Then
                                flg = 2
                            Else
                                CPT = 0
                                For K = 0 To 4
                                    If BX(K, X5) = SX2(K, Y5) Then
                                        CPT += 1
                                    End If
                                Next K
                                If CPT = 5 Then
                                    flg = 1
                                End If
                                X5 += 1
                            End If
                        End While
                        If flg = 0 Or flg = 2 Then
                            For X5 = 0 To 4
                                For V = 0 To NB
                                    If SX2(X5, Y5) = BQ(V) Then
                                        flg = 1
                                    End If
                                Next V
                            Next X5
                            If flg = 1 Then
                                '*** Elimination des quintes pairs,
                                '*** impairs, chevaux tous < 9 ou
                                '*** tous > 9
                                flgp = 0
                                flgi = 0
                                flgn = 0
                                For X5 = 0 To 4
                                    If SX2(X5, Y5) = 2 Or SX2(X5, Y5) = 4 Or SX2(X5, Y5) = 6 Or SX2(X5, Y5) = 8 Or SX2(X5, Y5) = 10 Or SX2(X5, Y5) = 12 Or SX2(X5, Y5) = 14 Or SX2(X5, Y5) = 16 Or SX2(X5, Y5) = 18 Or SX2(X5, Y5) = 20 Then
                                        flgp += 1
                                    Else
                                        flgi += 1
                                    End If
                                Next X5
                                If flgp = 5 Or flgi = 5 Then
                                    flgn = 1
                                End If
                                If SX2(0, Y5) <= 9 And SX2(1, Y5) <= 9 And SX2(2, Y5) <= 9 And SX2(3, Y5) <= 9 And SX2(4, Y5) <= 9 Then
                                    flgn = 1
                                End If
                                If SX2(0, Y5) >= 9 And SX2(1, Y5) >= 9 And SX2(2, Y5) >= 9 And SX2(3, Y5) >= 9 And SX2(4, Y5) >= 9 Then
                                    flgn = 1
                                End If
                                If flgn = 0 Then
                                    For X5 = 0 To 4
                                        BX(X5, TJ) = SX2(X5, Y5)
                                    Next X5
                                    TJ += 1
                                End If
                            End If
                        End If
                        Y5 += 1
                    End While
                Else
                    '*** Chargement des bases quinte si il reste 7 chevaux
                    If CPT = 6 Then
                        SX3(0, 0) = SX(0)
                        SX3(1, 0) = SX(1)
                        SX3(2, 0) = SX(2)
                        SX3(3, 0) = SX(3)
                        SX3(4, 0) = SX(4)
                        SX3(0, 1) = SX(0)
                        SX3(1, 1) = SX(1)
                        SX3(2, 1) = SX(2)
                        SX3(3, 1) = SX(3)
                        SX3(4, 1) = SX(5)
                        SX3(0, 2) = SX(0)
                        SX3(1, 2) = SX(1)
                        SX3(2, 2) = SX(2)
                        SX3(3, 2) = SX(5)
                        SX3(4, 2) = SX(6)
                        SX3(0, 3) = SX(0)
                        SX3(1, 3) = SX(1)
                        SX3(2, 3) = SX(3)
                        SX3(3, 3) = SX(4)
                        SX3(4, 3) = SX(5)
                        SX3(0, 4) = SX(0)
                        SX3(1, 4) = SX(2)
                        SX3(2, 4) = SX(3)
                        SX3(3, 4) = SX(4)
                        SX3(4, 4) = SX(5)
                        SX3(0, 5) = SX(1)
                        SX3(1, 5) = SX(2)
                        SX3(2, 5) = SX(3)
                        SX3(3, 5) = SX(4)
                        SX3(4, 5) = SX(5)
                        SX3(0, 6) = SX(0)
                        SX3(1, 6) = SX(1)
                        SX3(2, 6) = SX(2)
                        SX3(3, 6) = SX(3)
                        SX3(4, 6) = SX(6)
                        SX3(0, 7) = SX(0)
                        SX3(1, 7) = SX(1)
                        SX3(2, 7) = SX(2)
                        SX3(3, 7) = SX(4)
                        SX3(4, 7) = SX(6)
                        SX3(0, 8) = SX(0)
                        SX3(1, 8) = SX(1)
                        SX3(2, 8) = SX(3)
                        SX3(3, 8) = SX(4)
                        SX3(4, 8) = SX(6)
                        SX3(0, 9) = SX(0)
                        SX3(1, 9) = SX(2)
                        SX3(2, 9) = SX(3)
                        SX3(3, 9) = SX(4)
                        SX3(4, 9) = SX(6)
                        SX3(0, 10) = SX(1)
                        SX3(1, 10) = SX(2)
                        SX3(2, 10) = SX(3)
                        SX3(3, 10) = SX(4)
                        SX3(4, 10) = SX(6)
                        SX3(0, 11) = SX(0)
                        SX3(1, 11) = SX(1)
                        SX3(2, 11) = SX(2)
                        SX3(3, 11) = SX(5)
                        SX3(4, 11) = SX(6)
                        SX3(0, 12) = SX(0)
                        SX3(1, 12) = SX(1)
                        SX3(2, 12) = SX(3)
                        SX3(3, 12) = SX(5)
                        SX3(4, 12) = SX(6)
                        SX3(0, 13) = SX(0)
                        SX3(1, 13) = SX(2)
                        SX3(2, 13) = SX(3)
                        SX3(3, 13) = SX(5)
                        SX3(4, 13) = SX(6)
                        SX3(0, 14) = SX(1)
                        SX3(1, 14) = SX(2)
                        SX3(2, 14) = SX(3)
                        SX3(3, 14) = SX(5)
                        SX3(4, 14) = SX(6)
                        SX3(0, 15) = SX(0)
                        SX3(1, 15) = SX(1)
                        SX3(2, 15) = SX(4)
                        SX3(3, 15) = SX(5)
                        SX3(4, 15) = SX(6)
                        SX3(0, 16) = SX(0)
                        SX3(1, 16) = SX(2)
                        SX3(2, 16) = SX(4)
                        SX3(3, 16) = SX(5)
                        SX3(4, 16) = SX(6)
                        SX3(0, 17) = SX(1)
                        SX3(1, 17) = SX(2)
                        SX3(2, 17) = SX(4)
                        SX3(3, 17) = SX(5)
                        SX3(4, 17) = SX(6)
                        SX3(0, 18) = SX(0)
                        SX3(1, 18) = SX(3)
                        SX3(2, 18) = SX(4)
                        SX3(3, 18) = SX(5)
                        SX3(4, 18) = SX(6)
                        SX3(0, 19) = SX(1)
                        SX3(1, 19) = SX(3)
                        SX3(2, 19) = SX(4)
                        SX3(3, 19) = SX(5)
                        SX3(4, 19) = SX(6)
                        SX3(0, 20) = SX(2)
                        SX3(1, 20) = SX(3)
                        SX3(2, 20) = SX(4)
                        SX3(3, 20) = SX(5)
                        SX3(4, 20) = SX(6)
                        '*** tant que les 5 chevaux ne sont pas
                        '*** ne sont pas dans l'ordre
                        For Y5 = 0 To 20
                            For X5 = 0 To 4
                                For K = (X5 + 1) To 4
                                    If SX3(X5, Y5) <= SX3(K, Y5) Then
                                        X = SX3(K, Y5)
                                        SX3(K, Y5) = SX3(X5, Y5)
                                        SX3(X5, Y5) = X
                                    End If
                                Next K
                            Next X5
                        Next Y5
                        '*** tant que tous les quintes pas teste
                        '*** ou que quinte déja enregistree
                        Y5 = 0
                        While Y5 < 21
                            flg = 0
                            X5 = 0
                            While X5 < TJ And flg = 0
                                If BX(1, X5) = 0 Then
                                    flg = 2
                                Else
                                    CPT = 0
                                    For K = 0 To 4
                                        If BX(K, X5) = SX3(K, Y5) Then
                                            CPT += 1
                                        End If
                                    Next K
                                    If CPT = 5 Then
                                        flg = 1
                                    End If
                                    X5 = X5 + 1
                                End If
                            End While
                            If flg = 0 Or flg = 2 Then
                                For X5 = 0 To 4
                                    For V = 0 To NB
                                        If SX3(X5, Y5) = BQ(V) Then
                                            flg = 1
                                        End If
                                    Next V
                                Next X5
                                If flg = 1 Then
                                    '*** Elimination des quintes pairs,
                                    '*** impairs chevaux tous < 9 ou
                                    '*** tous > 9
                                    flgp = 0
                                    flgi = 0
                                    flgn = 0
                                    For X5 = 0 To 4
                                        If SX3(X5, Y5) = 2 Or SX3(X5, Y5) = 4 Or SX3(X5, Y5) = 6 Or SX3(X5, Y5) = 8 Or SX3(X5, Y5) = 10 Or SX3(X5, Y5) = 12 Or SX3(X5, Y5) = 14 Or SX3(X5, Y5) = 16 Or SX3(X5, Y5) = 18 Or SX3(X5, Y5) = 20 Then
                                            flgp += 1
                                        Else
                                            flgi += 1
                                        End If
                                    Next X5
                                    If flgp = 5 Or flgi = 5 Then
                                        flgn = 1
                                    End If
                                    If SX3(0, Y5) <= 9 And SX3(1, Y5) <= 9 And SX3(2, Y5) <= 9 And SX3(3, Y5) <= 9 And SX3(4, Y5) <= 9 Then
                                        flgn = 1
                                    End If
                                    If SX3(0, Y5) >= 9 And SX3(1, Y5) >= 9 And SX3(2, Y5) >= 9 And SX3(3, Y5) >= 9 And SX3(4, Y5) >= 9 Then
                                        flgn = 1
                                    End If
                                    If flgn = 0 Then
                                        For X5 = 0 To 4
                                            BX(X5, TJ) = SX3(X5, Y5)
                                        Next X5
                                        TJ += 1
                                    End If
                                End If
                            End If
                            Y5 += 1
                        End While
                    End If
                End If
            End If
            I += 1
        End While
        '*** Classement des chevaux suivant l'ordre de la selection
        I = 0
        ProgressBar2.Value = ProgressBar2.Maximum
        While I < 999 And BX(0, I) > 0
            For J = 0 To 4
                For K = 0 To 13
                    If BX(J, I) = B(K) Then
                        T1(J) = K           'POSITION DANS LA BASE
                        P1(J) = J           'POSITION DANS LA TABLE QUINTE
                    End If
                Next K
            Next J
            Do                                    'MISE EN ORDRE
                X = 0
                For N = 0 To 3
                    If T1(N) > T1(N + 1) Then
                        TX = T1(N)
                        PX = P1(N)
                        T1(N) = T1(N + 1)
                        P1(N) = P1(N + 1)
                        T1(N + 1) = TX
                        P1(N + 1) = PX
                        X += 1
                    End If
                Next N
                If X = 0 Then
                    Exit Do
                End If
            Loop
            QTN(0) = BX(P1(0), I)  'RECUPERATION DES NUMEROS DE
            QTN(1) = BX(P1(1), I)  'CHEVAUX PAR RAPPORT A LEUR
            QTN(2) = BX(P1(2), I)  'POSITION
            QTN(3) = BX(P1(3), I)
            QTN(4) = BX(P1(4), I)
            BX(0, I) = QTN(0)      'REECRITURE DANS LA BASE QUINTE
            BX(1, I) = QTN(1)      'SUIVANT L'ORDRE DETERMINE
            BX(2, I) = QTN(2)
            BX(3, I) = QTN(3)
            BX(4, I) = QTN(4)
            I += 1
        End While
        '*** Tri des quintes avant enregistrement dans les fichiers
        '*** mise en format sous forme de chaine de caractere
        I = 0
        While I < 999 And BX(0, I) > 0
            QT(I) = Format(BX(0, I), "00") & Format(BX(1, I), "00") & Format(BX(2, I), "00") & Format(BX(3, I), "00") & Format(BX(4, I), "00")
            I += 1
        End While
        '*** Formule de tri en parallele
        IM = I
        J = 0
        While J < IM
            I = 0
            QTM = "0000000000"
            While I < IM
                If QTM < QT(I) Then
                    QTM = QT(I)
                    RI = I
                End If
                I += 1
            End While
            QT(RI) = "0000000000"
            BXT(0, J) = BX(0, RI)
            BXT(1, J) = BX(1, RI)
            BXT(2, J) = BX(2, RI)
            BXT(3, J) = BX(3, RI)
            BXT(4, J) = BX(4, RI)
            J += 1
        End While
        '***
        '*** Sauvegarde des données quintes sous forme de fichiers
        '*** 
        On Error Resume Next
        File.Delete(DateCourse & ".REF")
        On Error Resume Next
        File.Delete(DateCourse & ".TXT")
        For I = 0 To 9
            If File.Exists(DateCourse & ".TR" & I) Then
                File.Delete(DateCourse & ".TR" & I)
            Else
                Exit For
            End If
        Next I
        MyFileW = My.Computer.FileSystem.OpenTextFileWriter(DateCourse & ".REF", False)
        MyFileW.WriteLine(B(0) & "," & B(1) & "," & B(2) & "," & B(3) & "," & B(4) & "," & B(5) & "," & B(6) & "," & B(7) & "," & B(8) & "," & B(9) & "," & B(10) & "," & B(11) & "," & B(12) & "," & B(13) & "," & BQ(0) & "," & BQ(1) & "," & BQ(2) & "," & ARQT(0) & "," & ARQT(1) & "," & ARQT(2) & "," & ARQT(3) & "," & ARQT(4))
        MyFileW.Close()
        I = 0
        BAS = 0
        TXTFileW = My.Computer.FileSystem.OpenTextFileWriter(DateCourse & ".TXT", False)
        TRIFileW = My.Computer.FileSystem.OpenTextFileWriter(DateCourse & ".TR0", False)
        Do
            TXTFileW.WriteLine(BXT(0, I) & "," & BXT(1, I) & "," & BXT(2, I) & "," & BXT(3, I) & "," & BXT(4, I))
            TRIFileW.WriteLine(BXT(0, I) & "," & BXT(1, I) & "," & BXT(2, I) & "," & BXT(3, I) & "," & BXT(4, I))
            If BXT(0, I) <> 0 Then
                BAS += 1
            Else
                Exit Do
            End If
            I += 1
        Loop Until I = 999
        TXTFileW.Close()
        TRIFileW.Close()
        '*** mise en forme de l'affichage
        SELMOINS.Enabled = False
        SELPLUS.Enabled = False
        Generation.Enabled = True
        INIT.Enabled = False
        COPY.Enabled = True
        Generation.Focus()
        WINQUINT.TabPages.Item("ELIMINATION").Enabled = True
        WINQUINT.TabPages.Item("VISUALISATION").Enabled = True
        Call ChargementListeQuintes()
        ProgressBar1.Visible = False
        ProgressBar2.Visible = False
    End Sub

    Private Sub ORDRE_Click(sender As Object, e As EventArgs) Handles ORDRE.Click
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer
        Dim x As Integer
        Dim n As Integer
        Dim tx As Integer
        Dim px As Integer
        Dim IM As Integer
        Dim QTM As Integer
        Dim RI As Integer
        Dim SCH As Integer
        Dim CptEL As Integer
        Dim Pos As Integer
        Dim Ok As Boolean
        Dim t1(4) As Integer
        Dim p1(4) As Integer

        If RES = 0 Then
            Exit Sub
        End If
        If ord = 0 Then
            For I = 0 To 13
                ChxTri(I).Text = Chevaux(I).Text
                ChxTri(I).Visible = True
            Next I
            ANNUL.Enabled = True
            CH1.Text = ""
            CH2.Text = ""
            CH3.Text = ""
            CH4.Text = ""
            CH5.Text = ""
            CH6.Text = ""
            CH7.Text = ""
            CH8.Text = ""
            CH9.Text = ""
            CH10.Text = ""
            CH11.Text = ""
            CH12.Text = ""
            CH13.Text = ""
            CH14.Text = ""
            WINQUINT.TabPages.Item("ELIMINATION").Enabled = False
            WINQUINT.TabPages.Item("VISUALISATION").Enabled = False
            BA1.Enabled = False
            BA2.Enabled = False
            BA3.Enabled = False
            SELEC.Enabled = False
            NBCOMBI.Enabled = False
            Generation.Enabled = False
            INIT.Enabled = False
            COPY.Enabled = False
            SELPLUS.Enabled = False
            SELMOINS.Enabled = False
            DatePlus.Enabled = False
            DateMoins.Enabled = False
            DateJour.Enabled = False
            NbBase1.Enabled = False
            NbBase2.Enabled = False
            NbBase3.Enabled = False
            ord = 1
            Exit Sub
        End If
        '*** Test de validité des chevaux saisis pour la mise en ordre
        CptEL = 0
        For I = 0 To 13
            SCH = Val(Chevaux(I).Text)
            If SCH > 0 Then
                CptEL += 1
            Else
                MsgBox("Donnée obligatoire pour procéder à une mise en ordre.", 48, "Erreur")
                Chevaux(I).Focus()
                Exit Sub
            End If
            Ok = True
            Pos = 1
            Call OrdreVerification()
            If Ok = False Then
                Chevaux(I).Focus()
                Exit Sub
            Else
                CH(I) = Val(Chevaux(I).Text)
            End If
        Next I

        '*** Test de cohérence des saisies
        If CptEL <> 14 Then
            MsgBox("14 chevaux sont nécessaires pour la mise en ordre.", 48, "Erreur")
            Chevaux(0).Focus()
            Exit Sub
        End If

        ANNUL.Enabled = True

        '*** Classement des chevaux suivant l'ordre saisi
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = RES
        ProgressBar1.Visible = True
        ProgressBar1.Value = ProgressBar1.Minimum

        '*** Chargement de la table temporaire
        For I = 0 To 999
            BX(0, I) = BXT(0, I)
            BX(1, I) = BXT(1, I)
            BX(2, I) = BXT(2, I)
            BX(3, I) = BXT(3, I)
            BX(4, I) = BXT(4, I)
        Next I
        I = 0
        While I < 999 And BX(0, I) > 0
            For J = 0 To 4
                For K = 0 To 13
                    If BX(J, I) = CH(K) Then
                        t1(J) = K                 'POSITION DANS LA BASE
                        p1(J) = J                 'POSITION DANS LA TABLE QUINTE
                    End If
                Next K
            Next J
            Do                                    'MISE EN ORDRE
                x = 0
                For n = 0 To 3
                    If t1(n) > t1(n + 1) Then
                        tx = t1(n)
                        px = p1(n)
                        t1(n) = t1(n + 1)
                        p1(n) = p1(n + 1)
                        t1(n + 1) = tx
                        p1(n + 1) = px
                        x = x + 1
                    End If
                Next n
                If x = 0 Then
                    Exit Do
                End If
            Loop
            QTN(0) = BX(p1(0), I)  'RECUPERATION DES NUMEROS DE
            QTN(1) = BX(p1(1), I)  'CHEVAUX PAR RAPPORT A LEUR
            QTN(2) = BX(p1(2), I)  'POSITION
            QTN(3) = BX(p1(3), I)
            QTN(4) = BX(p1(4), I)
            BX(0, I) = QTN(0)      'REECRITURE DANS LA BASE QUINTE
            BX(1, I) = QTN(1)      'SUIVANT L'ORDRE DETERMINE
            BX(2, I) = QTN(2)
            BX(3, I) = QTN(3)
            BX(4, I) = QTN(4)
            I += 1
            ProgressBar1.Value = I - 1
        End While
        '*** Tri des quintes avant enregistrement dans les fichiers
        '*** mise en format sous forme de chaine de caractere
        I = 0
        While I < 999 And BX(0, I) > 0
            QT(I) = Format(BX(0, I), "00") & Format(BX(1, I), "00") & Format(BX(2, I), "00") & Format(BX(3, I), "00") & Format(BX(4, I), "00")
            I += 1
        End While
        '*** Formule de tri en parallele
        IM = I
        J = 0
        While J < IM
            I = 0
            QTM = "0000000000"
            While I < IM
                If QTM < QT(I) Then
                    QTM = QT(I)
                    RI = I
                End If
                I += 1
            End While
            QT(RI) = "0000000000"
            BXT(0, J) = BX(0, RI)
            BXT(1, J) = BX(1, RI)
            BXT(2, J) = BX(2, RI)
            BXT(3, J) = BX(3, RI)
            BXT(4, J) = BX(4, RI)
            J += 1
        End While
        ProgressBar1.Value = ProgressBar1.Maximum

        '*** sauvegarde des données quintes dans le fichier de selection
        I = 0
        TRIFileW = My.Computer.FileSystem.OpenTextFileWriter(DateCourse & ".TR" & NumFile, False)
        While I < 999
            TRIFileW.WriteLine(BXT(0, I) & "," & BXT(1, I) & "," & BXT(2, I) & "," & BXT(3, I) & "," & BXT(4, I))
            I += 1
        End While
        TRIFileW.Close()

        ProgressBar1.Visible = False

        ORDRE.Enabled = False
        CH1.Text = B(0)
        CH2.Text = B(1)
        CH3.Text = B(2)
        CH4.Text = B(3)
        CH5.Text = B(4)
        CH6.Text = B(5)
        CH7.Text = B(6)
        CH8.Text = B(7)
        CH9.Text = B(8)
        CH10.Text = B(9)
        CH11.Text = B(10)
        CH12.Text = B(11)
        CH13.Text = B(12)
        CH14.Text = B(13)
        WINQUINT.TabPages.Item("SYNTHESE").Enabled = True
        WINQUINT.TabPages.Item("VISUALISATION").Enabled = True
        BA1.Enabled = True
        BA2.Enabled = True
        BA3.Enabled = True
        SELEC.Enabled = True
        NBCOMBI.Enabled = True
        INIT.Enabled = True
        COPY.Enabled = True
        Generation.Enabled = True
        SELPLUS.Enabled = True
        SELMOINS.Enabled = True
        DatePlus.Enabled = True
        DateMoins.Enabled = True
        DateJour.Enabled = True
        NbBase1.Enabled = True
        NbBase2.Enabled = True
        NbBase3.Enabled = True
        For I = 0 To 13
            ChxTri(I).Visible = False
        Next I
        Generation.Focus()
    End Sub

    Private Sub OrdreVerification()
        Dim I As Integer
        If SCH = 0 Then
            Ok = True
            Exit Sub
        End If
        If SCH < 1 Or SCH > 20 Then
            MsgBox("Donnée strictement numérique" + vbCrLf + "comprise entre 1 et 20.", 48, "Erreur")
            Ok = False
        Else
            While I <> 14
                If SCH = B(I) Then
                    Ok = False
                    I = 13
                End If
                I += 1
            End While
            If Ok = False Then
                Ok = True
            Else
                Ok = False
            End If
            For I = 0 To Pos - 2
                If SCH = CH(I) Then
                    Ok = False
                    I = Pos - 1
                End If
            Next I
            If Ok = False Then
                MsgBox("Les données doivent être inclues dans la synthèse" + vbCrLf + "et strictement différentes.", 48, "Erreur")
            End If
        End If
    End Sub

    Private Sub ANNUL_Click(sender As Object, e As EventArgs) Handles ANNUL.Click
        Dim I As Integer
        CH1.Text = B(0)
        CH2.Text = B(1)
        CH3.Text = B(2)
        CH4.Text = B(3)
        CH5.Text = B(4)
        CH6.Text = B(5)
        CH7.Text = B(6)
        CH8.Text = B(7)
        CH9.Text = B(8)
        CH10.Text = B(9)
        CH11.Text = B(10)
        CH12.Text = B(11)
        CH13.Text = B(12)
        CH14.Text = B(13)
        WINQUINT.TabPages.Item("ELIMINATION").Enabled = True
        WINQUINT.TabPages.Item("VISUALISATION").Enabled = True
        BA1.Enabled = True
        BA2.Enabled = True
        BA3.Enabled = True
        SELEC.Enabled = True
        NBCOMBI.Enabled = True
        INIT.Enabled = True
        COPY.Enabled = True
        Generation.Enabled = True
        SELPLUS.Enabled = True
        SELMOINS.Enabled = True
        DatePlus.Enabled = True
        DateMoins.Enabled = True
        DateJour.Enabled = True
        ORDRE.Enabled = True
        NbBase1.Enabled = True
        NbBase2.Enabled = True
        NbBase3.Enabled = True
        ANNUL.Enabled = False
        For I = 0 To 13
            ChxTri(I).Visible = False
        Next I
        Generation.Focus()
        ord = 0
    End Sub

    Private Sub INIT_Click(sender As Object, e As EventArgs) Handles INIT.Click
        Dim I As Integer
        If MsgBox("Confirmez-vous la réinitialisation de la sélection ?", 48, "REGENERER LA SELECTION") Then
            For I = 0 To 9
                On Error Resume Next
                File.Delete(DateCourse & ".TR" & I)
            Next I
            File.Copy(DateCourse & ".TXT", DateCourse & ".TR0")
            NumFile = 0
            MaxFile = NumFile
            Call ChargementListeQuintes()
        End If
    End Sub

    Private Sub COPY_Click(sender As Object, e As EventArgs) Handles COPY.Click
        Dim Reponse As Integer
        Dim I As Integer
        Reponse = File.Exists(DateCourse & ".TR" & NumFile + 1)
        If Reponse Then
            MsgBox("Voulez-vous écraser la sélection N°" & NumFile & " et les suivantes ?", 48, "SAUVEGARDER CETTE SELECTION")
            For I = NumFile To 9
                On Error Resume Next
                File.Delete(DateCourse & ".TR" & I)
            Next I
            If NumFile = 0 Then
                File.Copy(DateCourse & ".TXT", DateCourse & ".TR0")
            End If
        End If
        File.Copy(DateCourse & ".TR" & NumFile, DateCourse & ".TR" & NumFile + 1)
        SELMOINS.Enabled = True
        If NumFile >= 9 Then
            COPY.Enabled = False
            NumFile = 9
        Else
            NumFile += 1
        End If
        MaxFile = NumFile
        Call ChargementListeQuintes()
    End Sub

    Private Sub ChevauxVerifications()
        Dim I As Integer
        If Not IsNumeric(SCH) Or SCH < 1 Or SCH > 20 Then
            MsgBox("Donnée obligatoire et numérique" + vbCrLf + "comprise entre 1 et 20.", vbExclamation, "Erreur")
            Ok = False
        Else
            I = 0
            While I < (Pos - 1) And Pos > 1
                If SCH = B(I) Then
                    Ok = False
                    MsgBox("Les données doivent être strictement différentes.", vbExclamation, "Erreur")
                    I = Pos - 1
                End If
                I += 1
            End While
        End If
    End Sub

    Private Sub BasesVerifications()
        Dim I As Integer
        If Not IsNumeric(SCH) Or SCH < 1 Or SCH > 20 Then
            MsgBox("Donnée obligatoire et numérique" + vbCrLf + "comprise entre 1 et 20.", vbExclamation, "Erreur")
            Ok = False
        Else
            For I = 0 To 13
                If SCH = B(I) Then
                    Ok = False
                End If
            Next I
            If Ok = True Then
                MsgBox("Les bases doivent faire parties de la synthèse.", vbExclamation, "Erreur")
                Ok = False
            Else
                Ok = True
            End If
        End If
    End Sub

    Sub Initialisation()
        Dim I As Integer
        Dim J As Integer
        '*** Initialisation des tables synthese, bases et bases16
        For I = 0 To 13
            B(I) = 0
        Next I
        For I = 0 To 2
            BQ(I) = 0
        Next I
        For I = 0 To 4
            For J = 1 To 16
                T(I, J) = 0
            Next J
        Next I
        For I = 0 To 4
            ARQT(I) = 0
            For J = 0 To 999
                BXT(I, J) = 0
            Next J
        Next I
        NbBase1.Checked = True
        BA1.Visible = True
        BA2.Visible = False
        BA3.Visible = False
        BA1.Text = ""
        BA2.Text = ""
        BA3.Text = ""
        INIT.Enabled = False
        COPY.Enabled = False
        NBCOMBI.Text = ""
        SELEC.Text = ""
        CPTELI.Text = ""
        CPTRES.Text = ""
        BA1.Text = ""
        BA2.Text = ""
        BA3.Text = ""
        QTA1.Text = ""
        QTA2.Text = ""
        QTA3.Text = ""
        QTA4.Text = ""
        QTA5.Text = ""
        NBPAGE.Text = ""
        '***
        '*** Chargement des tableaux 14 chevaux
        '***
        Chevaux = New ArrayList()
        Chevaux.Add(CH1)
        Chevaux.Add(CH2)
        Chevaux.Add(CH3)
        Chevaux.Add(CH4)
        Chevaux.Add(CH5)
        Chevaux.Add(CH6)
        Chevaux.Add(CH7)
        Chevaux.Add(CH8)
        Chevaux.Add(CH9)
        Chevaux.Add(CH10)
        Chevaux.Add(CH11)
        Chevaux.Add(CH12)
        Chevaux.Add(CH13)
        Chevaux.Add(CH14)
        '***
        ChxTri = New ArrayList()
        ChxTri.Add(TR1)
        ChxTri.Add(TR2)
        ChxTri.Add(TR3)
        ChxTri.Add(TR4)
        ChxTri.Add(TR5)
        ChxTri.Add(TR6)
        ChxTri.Add(TR7)
        ChxTri.Add(TR8)
        ChxTri.Add(TR9)
        ChxTri.Add(TR10)
        ChxTri.Add(TR11)
        ChxTri.Add(TR12)
        ChxTri.Add(TR13)
        ChxTri.Add(TR14)
        '***
        ChxSyn = New ArrayList()
        ChxSyn.Add(SY1)
        ChxSyn.Add(SY2)
        ChxSyn.Add(SY3)
        ChxSyn.Add(SY4)
        ChxSyn.Add(SY5)
        ChxSyn.Add(SY6)
        ChxSyn.Add(SY7)
        ChxSyn.Add(SY8)
        ChxSyn.Add(SY9)
        ChxSyn.Add(SY10)
        ChxSyn.Add(SY11)
        ChxSyn.Add(SY12)
        ChxSyn.Add(SY13)
        ChxSyn.Add(SY14)
        '***
        ChxElim = New ArrayList()
        ChxElim.Add(EL1)
        ChxElim.Add(EL2)
        ChxElim.Add(EL3)
        ChxElim.Add(EL4)
        ChxElim.Add(EL5)
        ChxElim.Add(EL6)
        ChxElim.Add(EL7)
        ChxElim.Add(EL8)
        ChxElim.Add(EL9)
        ChxElim.Add(EL10)
        ChxElim.Add(EL11)
        ChxElim.Add(EL12)
        ChxElim.Add(EL13)
        ChxElim.Add(EL14)
        '***
        Quintes = New ArrayList()
        Quintes.Add(QUINTE01)
        Quintes.Add(QUINTE02)
        Quintes.Add(QUINTE03)
        Quintes.Add(QUINTE04)
        Quintes.Add(QUINTE05)
        Quintes.Add(QUINTE06)
        Quintes.Add(QUINTE07)
        Quintes.Add(QUINTE08)
        Quintes.Add(QUINTE09)
        Quintes.Add(QUINTE10)
        Quintes.Add(QUINTE11)
        Quintes.Add(QUINTE12)
        Quintes.Add(QUINTE13)
        Quintes.Add(QUINTE14)
        Quintes.Add(QUINTE15)
        Quintes.Add(QUINTE16)
        Quintes.Add(QUINTE17)
        Quintes.Add(QUINTE18)
        Quintes.Add(QUINTE19)
        Quintes.Add(QUINTE20)
        Quintes.Add(QUINTE21)
        Quintes.Add(QUINTE22)
        Quintes.Add(QUINTE23)
        Quintes.Add(QUINTE24)
        Quintes.Add(QUINTE25)
        Quintes.Add(QUINTE26)
        Quintes.Add(QUINTE27)
        Quintes.Add(QUINTE28)
        Quintes.Add(QUINTE29)
        Quintes.Add(QUINTE30)
        Quintes.Add(QUINTE31)
        Quintes.Add(QUINTE32)
        Quintes.Add(QUINTE33)
        Quintes.Add(QUINTE34)
        Quintes.Add(QUINTE35)
        Quintes.Add(QUINTE36)
        Quintes.Add(QUINTE37)
        Quintes.Add(QUINTE38)
        Quintes.Add(QUINTE39)
        Quintes.Add(QUINTE40)
        Quintes.Add(QUINTE41)
        Quintes.Add(QUINTE42)
        Quintes.Add(QUINTE43)
        Quintes.Add(QUINTE44)
        Quintes.Add(QUINTE45)
        Quintes.Add(QUINTE46)
        Quintes.Add(QUINTE47)
        Quintes.Add(QUINTE48)
        Quintes.Add(QUINTE49)
        Quintes.Add(QUINTE50)
        Quintes.Add(QUINTE51)
        Quintes.Add(QUINTE52)
        Quintes.Add(QUINTE53)
        Quintes.Add(QUINTE54)
        Quintes.Add(QUINTE55)
        Quintes.Add(QUINTE56)
        Quintes.Add(QUINTE57)
        Quintes.Add(QUINTE58)
        Quintes.Add(QUINTE59)
        Quintes.Add(QUINTE60)
        Quintes.Add(QUINTE61)
        Quintes.Add(QUINTE62)
        Quintes.Add(QUINTE63)
        Quintes.Add(QUINTE64)
        Quintes.Add(QUINTE65)
        Quintes.Add(QUINTE66)
        Quintes.Add(QUINTE67)
        Quintes.Add(QUINTE68)
        Quintes.Add(QUINTE69)
        Quintes.Add(QUINTE70)
        '***
        For I = 0 To 13
            Chevaux(I).Text = ""
            ChxElim(I).Text = ""
            ChxSyn(I).Text = ""
            ChxTri(I).Text = ""
        Next I
    End Sub

    Sub ChargementGrilleSynthese()
        'constitution de la grille des SYNTHESES 5 et 3
        'base 5 ligne 1
        B5(0, 0) = Val(CH7.Text)
        B5(1, 0) = Val(CH8.Text)
        B5(2, 0) = Val(CH9.Text)
        B5(3, 0) = Val(CH10.Text)
        B5(4, 0) = Val(CH11.Text)
        'base 3 ligne 1
        B3(0, 0) = Val(CH12.Text)
        B3(1, 0) = Val(CH13.Text)
        B3(2, 0) = Val(CH14.Text)
        'base 5 ligne 2
        B5(0, 1) = Val(CH11.Text)
        B5(1, 1) = Val(CH4.Text)
        B5(2, 1) = Val(CH10.Text)
        B5(3, 1) = Val(CH5.Text)
        B5(4, 1) = Val(CH9.Text)
        'base 3 ligne 2
        B3(0, 1) = Val(CH6.Text)
        B3(1, 1) = Val(CH8.Text)
        B3(2, 1) = Val(CH7.Text)
        'base 5 ligne 3
        B5(0, 2) = Val(CH9.Text)
        B5(1, 2) = Val(CH2.Text)
        B5(2, 2) = Val(CH3.Text)
        B5(3, 2) = Val(CH12.Text)
        B5(4, 2) = Val(CH10.Text)
        'base 3 ligne 3
        B3(0, 2) = Val(CH3.Text)
        B3(1, 2) = Val(CH4.Text)
        B3(2, 2) = Val(CH11.Text)
        'base 5 ligne 4
        B5(0, 3) = Val(CH10.Text)
        B5(1, 3) = Val(CH1.Text)
        B5(2, 3) = Val(CH12.Text)
        B5(3, 3) = Val(CH6.Text)
        B5(4, 3) = Val(CH5.Text)
        'base 3 ligne 4
        B3(0, 3) = Val(CH13.Text)
        B3(1, 3) = Val(CH2.Text)
        B3(2, 3) = Val(CH9.Text)
        'base 5 ligne 5
        B5(0, 4) = Val(CH5.Text)
        B5(1, 4) = Val(CH14.Text)
        B5(2, 4) = Val(CH6.Text)
        B5(3, 4) = Val(CH3.Text)
        B5(4, 4) = Val(CH12.Text)
        'base 3 ligne 5
        B3(0, 4) = Val(CH8.Text)
        B3(1, 4) = Val(CH1.Text)
        B3(2, 4) = Val(CH10.Text)
        'base 5 ligne 6
        B5(0, 5) = Val(CH12.Text)
        B5(1, 5) = Val(CH7.Text)
        B5(2, 5) = Val(CH3.Text)
        B5(3, 5) = Val(CH13.Text)
        B5(4, 5) = Val(CH6.Text)
        'base 3 ligne 6
        B3(0, 5) = Val(CH4.Text)
        B3(1, 5) = Val(CH14.Text)
        B3(2, 5) = Val(CH5.Text)
        'base 5 ligne 7
        B5(0, 6) = Val(CH6.Text)
        B5(1, 6) = Val(CH11.Text)
        B5(2, 6) = Val(CH13.Text)
        B5(3, 6) = Val(CH8.Text)
        B5(4, 6) = Val(CH3.Text)
        'base 3 ligne 7
        B3(0, 6) = Val(CH2.Text)
        B3(1, 6) = Val(CH7.Text)
        B3(2, 6) = Val(CH12.Text)
        'base 5 ligne 8
        B5(0, 7) = Val(CH3.Text)
        B5(1, 7) = Val(CH9.Text)
        B5(2, 7) = Val(CH8.Text)
        B5(3, 7) = Val(CH4.Text)
        B5(4, 7) = Val(CH13.Text)
        'base 3 ligne 8
        B3(0, 7) = Val(CH1.Text)
        B3(1, 7) = Val(CH11.Text)
        B3(2, 7) = Val(CH6.Text)
        'base 5 ligne 9
        B5(0, 8) = Val(CH13.Text)
        B5(1, 8) = Val(CH10.Text)
        B5(2, 8) = Val(CH4.Text)
        B5(3, 8) = Val(CH2.Text)
        B5(4, 8) = Val(CH8.Text)
        'base 3 ligne 9
        B3(0, 8) = Val(CH14.Text)
        B3(1, 8) = Val(CH9.Text)
        B3(2, 8) = Val(CH3.Text)
        'base 5 ligne 10
        B5(0, 9) = Val(CH8.Text)
        B5(1, 9) = Val(CH5.Text)
        B5(2, 9) = Val(CH2.Text)
        B5(3, 9) = Val(CH1.Text)
        B5(4, 9) = Val(CH4.Text)
        'base 3 ligne 10
        B3(0, 9) = Val(CH7.Text)
        B3(1, 9) = Val(CH10.Text)
        B3(2, 9) = Val(CH13.Text)
        'base 5 ligne 11
        B5(0, 10) = Val(CH4.Text)
        B5(1, 10) = Val(CH12.Text)
        B5(2, 10) = Val(CH1.Text)
        B5(3, 10) = Val(CH14.Text)
        B5(4, 10) = Val(CH2.Text)
        'base 3 ligne 11
        B3(0, 10) = Val(CH11.Text)
        B3(1, 10) = Val(CH5.Text)
        B3(2, 10) = Val(CH8.Text)
        'base 5 ligne 12
        B5(0, 11) = Val(CH2.Text)
        B5(1, 11) = Val(CH6.Text)
        B5(2, 11) = Val(CH14.Text)
        B5(3, 11) = Val(CH7.Text)
        B5(4, 11) = Val(CH1.Text)
        'base 3 ligne 12
        B3(0, 11) = Val(CH9.Text)
        B3(1, 11) = Val(CH12.Text)
        B3(2, 11) = Val(CH4.Text)
        'base 5 ligne 13
        B5(0, 12) = Val(CH1.Text)
        B5(1, 12) = Val(CH3.Text)
        B5(2, 12) = Val(CH7.Text)
        B5(3, 12) = Val(CH11.Text)
        B5(4, 12) = Val(CH14.Text)
        'base 3 ligne 13
        B3(0, 12) = Val(CH10.Text)
        B3(1, 12) = Val(CH6.Text)
        B3(2, 12) = Val(CH2.Text)
        'base 5 ligne 14
        B5(0, 13) = Val(CH14.Text)
        B5(1, 13) = Val(CH13.Text)
        B5(2, 13) = Val(CH11.Text)
        B5(3, 13) = Val(CH9.Text)
        B5(4, 13) = Val(CH7.Text)
        'base 3 ligne 14
        B3(0, 13) = Val(CH5.Text)
        B3(1, 13) = Val(CH3.Text)
        B3(2, 13) = Val(CH1.Text)
    End Sub

    Sub ChargementListeQuintes()
        Dim I As Integer
        Dim J As Integer
        Dim QTA1 As String
        Dim QTA2 As String
        Dim QTA3 As String
        Dim QTA4 As String
        Dim QTA5 As String
        If SELECFlag = False Then
            For I = 0 To 9
                If Not File.Exists(DateCourse & ".TR" & I) Then
                    Exit For
                End If
            Next I
            If I > 0 Then
                NumFile = I - 1
            Else
                NumFile = 0
            End If
            MaxFile = NumFile
        End If
        TRIFile = My.Computer.FileSystem.OpenTextFileReader(DateCourse & ".TR" & NumFile)
        I = 0
        CPTC = 0
        Do
            LQ = Split(TRIFile.ReadLine(), ",")
            If LQ(0) = "0" Then Exit Do
            For J = 0 To 4
                BXT(J, I) = LQ(J)
            Next J
            CPTC += 1
            I += 1
        Loop Until TRIFile.EndOfStream
        TRIFile.Close()
        '***
        RES = CPTC
        INIT.Enabled = True
        COPY.Enabled = True
        NBCOMBI.Text = CPTC & " Combinaisons"
        If ARQT(0) > 0 Then
            QTA1 = ARQT(0)
            QTA2 = ARQT(1)
            QTA3 = ARQT(2)
            QTA4 = ARQT(3)
            QTA5 = ARQT(4)
        Else
            QTA1 = ""
            QTA2 = ""
            QTA3 = ""
            QTA4 = ""
            QTA5 = ""
        End If
        NBQ.Text = ""
        NB4.Text = ""
        NB3.Text = ""
        For I = 0 To 13
            ChxSyn(I).Text = Chevaux(I).Text
            ChxSyn(I).BackColor = Color.Gold
            ChxSyn(I).Visible = True
        Next I
        '***
        '*** Mise en forme de l'affichage
        '*** Gestion de la navigation SELEC
        '***
        If MaxFile > NumFile Then
            SELPLUS.Enabled = True
        Else
            SELPLUS.Enabled = False
        End If
        If NumFile = 0 Then
            SELMOINS.Enabled = False
        Else
            SELMOINS.Enabled = True
        End If
        ProgressBar1.Visible = False
        ProgressBar2.Visible = False
        WINQUINT.TabPages.Item("ELIMINATION").Enabled = True
        WINQUINT.TabPages.Item("VISUALISATION").Enabled = True
        '***
        Call ChargementVisualisation()
        '***
        SELEC.Text = "N° " & NumFile
        SELECFlag = True
        '***
        DateJour.Focus()
    End Sub

    '***
    '*** ELIMINATION
    '***

    Private Sub ELIMIN_Click(sender As Object, e As EventArgs) Handles ELIMIN.Click
        Dim I As Integer
        Dim J As Integer
        Dim CptEL As Integer
        Dim NBS As Integer

        '*** Initialisation de la table des éliminations
        For I = 0 To 13
            EL(I) = 0
        Next I
        CptEL = 0
        '*** Test de validité des données saisies
        For I = 0 To 13
            SCH = Val(ChxElim(I).Text)
            If SCH > 0 Then
                CptEL += 1
            ElseIf I = 0 Then
                MsgBox("donnée obligatoire pour procéder à une élimination.", 48, "Erreur")
                ChxElim(I).Focus()
                Exit Sub
            End If
            Ok = True
            Pos = I + 1
            Call EliminationVerification()
            If Ok = False Then
                ChxElim(I).Focus()
                Exit Sub
            End If
        Next I

        '*** Aiguillage selon le nombre de chevaux saisis
        NBS = 0
        For I = 0 To 13
            If Val(ChxElim(I).Text) > 0 Then
                EL(I) = Val(ChxElim(I).Text)
                NBS += 1
            Else
                I = 13
            End If
        Next I

        '*** Test de cohérence des saisies
        If NBS <> CptEL Then
            MsgBox("Veuillez saisir sans cases a blancs S.V.P.", 48, "Erreur")
            EL1.Focus()
            Exit Sub
        End If

        'sauvegarde avant elimination
        I = 0
        While I < 999
            BXS(0, I) = BXT(0, I)
            BXS(1, I) = BXT(1, I)
            BXS(2, I) = BXT(2, I)
            BXS(3, I) = BXT(3, I)
            BXS(4, I) = BXT(4, I)
            I += 1
        End While
        FLGELI = 1
        ELIS = ELI
        RESS = RES
        ELI = 0
        RES = 0
        I = 0

        '*** Elimination pour moins de 6 chevaux saisis
        If NBS < 6 And NBS > 0 Then
            While BXT(0, I) > 0
                EG = 0
                J = 0
                While J < NBS
                    If EL(J) = BXT(0, I) Or EL(J) = BXT(1, I) Or EL(J) = BXT(2, I) Or EL(J) = BXT(3, I) Or EL(J) = BXT(4, I) Then
                        EG += 1
                    End If
                    J += 1
                End While
                If EG = NBS Then    'TOUS LES CHEVAUX SELECTIONNES SONT
                    For J = 0 To 4   'PRESENT DANS LA COMBINAISON QUINTE
                        BXT(J, I) = 0
                    Next J
                    ELI += 1
                End If
                I += 1
            End While
        End If

        '*** Elimination pour plus de 5 chevaux saisis
        If NBS > 5 Then
            While BXT(0, I) > 0
                EG = 0
                J = 0
                While J < NBS
                    If EL(J) = BXT(0, I) Or EL(J) = BXT(1, I) Or EL(J) = BXT(2, I) Or EL(J) = BXT(3, I) Or EL(J) = BXT(4, I) Then
                        EG += 1
                    End If
                    J += 1
                End While
                If EG = 5 Then      'LA COMBINAISON QUINTE EST PRESENTE
                    For J = 0 To 4   'DANS LA LISTE DES CHEVAUX A ELIMINER
                        BXT(J, I) = 0
                    Next J
                    ELI += 1
                End If
                I += 1
            End While
        End If

        '*** Remise en forme de la liste des combinaisons
        '*** avec elimination des combinaisons à zéro
        I = 0
        J = 0
        While I < 999
            If BXT(0, I) > 0 Then
                BX(0, J) = BXT(0, I)
                BX(1, J) = BXT(1, I)
                BX(2, J) = BXT(2, I)
                BX(3, J) = BXT(3, I)
                BX(4, J) = BXT(4, I)
                J += 1
                RES = J
            End If
            I += 1
        End While

        '*** Mise à zéro des combinaisons non utilisées
        For I = J To 999
            BX(0, I) = 0
            BX(1, I) = 0
            BX(2, I) = 0
            BX(3, I) = 0
            BX(4, I) = 0
        Next I

        '*** Rechargement de la liste des combinaisons triées
        For I = 0 To 999
            BXT(0, I) = BX(0, I)
            BXT(1, I) = BX(1, I)
            BXT(2, I) = BX(2, I)
            BXT(3, I) = BX(3, I)
            BXT(4, I) = BX(4, I)
        Next I

        '*** Enregistrements des combinaisons dans le fichier en cours
        I = 0
        On Error Resume Next
        File.Delete(DateCourse & ".TR" & NumFile)
        MyFileW = My.Computer.FileSystem.OpenTextFileWriter(DateCourse & ".TR" & NumFile, False)
        While I < 999
            MyFileW.WriteLine(BXT(0, I) & "," & BXT(1, I) & "," & BXT(2, I) & "," & BXT(3, I) & "," & BXT(4, I))
            I += 1
        End While
        MyFileW.Close()
        CPTELI.Text = ELI & " - Eliminées"
        CPTRES.Text = RES & " - Restantes"
        For I = 0 To 13
            ChxElim(I).Text = ""
        Next I
        Call ChargementVisualisation()
        EL1.Focus()
    End Sub

    Private Sub ANNULA_Click(sender As Object, e As EventArgs) Handles ANNULA.Click
        Dim I As Integer
        '*** Retablissement de la liste des combinaisons
        '*** avant elimination
        If FLGELI = 1 Then
            I = 0
            While I < 999
                BXT(0, I) = BXS(0, I)
                BXT(1, I) = BXS(1, I)
                BXT(2, I) = BXS(2, I)
                BXT(3, I) = BXS(3, I)
                BXT(4, I) = BXS(4, I)
                I += 1
            End While
            RES = RESS
            ELI = ELIS
            CPTELI.Text = ELI & " - Eliminées"
            CPTRES.Text = RES & " - Restantes"
        End If
    End Sub

    Private Sub EliminationQuintes()
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer

        '*** Elimination des combinaisons hors bonus 3
        I = 0
        Do While BXT(0, I) > 0
            EG = 0
            J = 0
            Do While J < 14
                For K = 0 To 4
                    If EL(J) = BXT(K, I) Then
                        EG += 1
                    End If
                Next K
                J += 1
            Loop
            If EG = 5 Then      '*** Mise à zéro des combinaisons sans valeur
                For J = 0 To 4
                    BXT(J, I) = 0
                Next J
                ELI += 1
            End If
            I += 1
        Loop
        '*** Rechargement de la table des Quintés
        '*** avec suppression des combinaisons sans valeurs
        I = 0
        J = 0
        Do While I < 999
            If BXT(0, I) > 0 Then
                For K = 0 To 4
                    BX(K, J) = BXT(K, I)
                Next K
                J += 1
            End If
            I += 1
        Loop
        RES = J
        '*** Mise à zéro des combinaisons non utilisées
        For I = J To 999
            For K = 0 To 4
                BX(K, I) = 0
            Next K
        Next I
        '*** Rechargement de la liste des combinaisons
        For I = 0 To 999
            For K = 0 To 4
                BXT(K, I) = BX(K, I)
            Next K
        Next I
    End Sub

    Private Sub EliminationVerification()
        Dim I As Integer
        If SCH = 0 Then
            Ok = True
            Exit Sub
        End If
        If SCH < 1 Or SCH > 20 Then
            MsgBox("Donnée strictement numérique" + vbCrLf + "comprise entre 1 et 20.", 48, "Erreur")
            Ok = False
        Else
            While I <> 14
                If SCH = B(I) Then
                    Ok = False
                    I = 13
                End If
                I += 1
            End While
            If Ok = False Then
                Ok = True
            Else
                Ok = False
            End If
            For I = 0 To Pos - 2
                If SCH = EL(I) Then
                    Ok = False
                    I = Pos - 1
                End If
            Next I
            If Ok = False Then
                MsgBox("Les données doivent être inclues dans la synthèse" + vbCrLf + "et strictement différentes.", 48, "Erreur")
            End If
        End If
    End Sub

    Private Sub ArriveeQuinte_Click(sender As Object, e As EventArgs) Handles ArriveeQuinte.Click
        Dim CPTQ As Integer
        Dim FlgQ As Integer
        Dim CPT4 As Integer
        Dim CPT3 As Integer
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer

        '*** Test de validite des données
        SCH = Val(QTA1.Text)
        Ok = True
        Pos = 1
        Call ArriveeVerification()
        If Ok = False Then
            QTA1.Focus()
            Exit Sub
        Else
            ARQT(0) = Val(QTA1.Text)
        End If
        SCH = Val(QTA2.Text)
        Ok = True
        Pos = 2
        Call ArriveeVerification()
        If Ok = False Then
            QTA2.Focus()
            Exit Sub
        Else
            ARQT(1) = Val(QTA2.Text)
        End If
        SCH = Val(QTA3.Text)
        Ok = True
        Pos = 3
        Call ArriveeVerification()
        If Ok = False Then
            QTA3.Focus()
            Exit Sub
        Else
            ARQT(2) = Val(QTA3.Text)
        End If
        SCH = Val(QTA4.Text)
        Ok = True
        Pos = 4
        Call ArriveeVerification()
        If Ok = False Then
            QTA4.Focus()
            Exit Sub
        Else
            ARQT(3) = Val(QTA4.Text)
        End If
        SCH = Val(QTA5.Text)
        Ok = True
        Pos = 5
        Call ArriveeVerification()
        If Ok = False Then
            QTA5.Focus()
            Exit Sub
        Else
            ARQT(4) = Val(QTA5.Text)
        End If
        '*** Elimination des combinaisons sans valeur
        ELI = 0
        J = 0
        For I = 0 To 13
            If B(I) <> ARQT(0) Then
                EL(J) = B(I)
                J += 1
            End If
        Next I
        If J = 13 Then
            EL(13) = 0
        End If
        Call EliminationQuintes()
        J = 0
        For I = 0 To 13
            If B(I) <> ARQT(1) Then
                EL(J) = B(I)
                J += 1
            End If
        Next I
        If J = 13 Then
            EL(13) = 0
        End If
        Call EliminationQuintes()
        J = 0
        For I = 0 To 13
            If B(I) <> ARQT(2) Then
                EL(J) = B(I)
                J += 1
            End If
        Next I
        If J = 13 Then
            EL(13) = 0
        End If
        Call EliminationQuintes()

        '*** Comptabilisation des combinaisons gagnantes
        For I = 0 To RES
            FlgQ = 0
            For J = 0 To 4
                For K = 0 To 4
                    If BXT(K, I) = ARQT(J) Then
                        FlgQ += 1
                        K = 4
                    End If
                Next K
                If J = 3 And FlgQ = 3 Then
                    J = 4
                End If
            Next J
            Select Case FlgQ
                Case 3
                    CPT3 += 1
                Case 4
                    CPT4 += 1
                Case 5
                    CPTQ = 1
            End Select
        Next I
        For I = 0 To 13
            ChxSyn(I).Visible = True
            ChxSyn(I).Text = B(I)
            If B(I) = ARQT(0) Or B(I) = ARQT(1) Or B(I) = ARQT(2) Or B(I) = ARQT(3) Or B(I) = ARQT(4) Then
                ChxSyn(I).BackColor = Color.Aqua
            Else
                ChxSyn(I).BackColor = Color.Gold
            End If
        Next I
        NBQ.Text = " - " & CPTQ & " - "
        NB4.Text = " - " & CPT4 & " - "
        NB3.Text = " - " & CPT3 & " - "
        CPTELI.Text = ELI & " - Eliminées"
        CPTRES.Text = RES & " - Restantes"

        '*** sauvegarde de l'arrivée dans le fichier références
        MyFileW = My.Computer.FileSystem.OpenTextFileWriter(DateCourse & ".REF", False)
        MyFileW.WriteLine(B(0) & "," & B(1) & "," & B(2) & "," & B(3) & "," & B(4) & "," & B(5) & "," & B(6) & "," & B(7) & "," & B(8) & "," & B(9) & "," & B(10) & "," & B(11) & "," & B(12) & "," & B(13) & "," & BQ(0) & "," & BQ(1) & "," & BQ(2) & "," & ARQT(0) & "," & ARQT(1) & "," & ARQT(2) & "," & ARQT(3) & "," & ARQT(4))
        MyFileW.Close()
        For I = 0 To 13
            ChxElim(I).Enabled = False
        Next I
        ELIMIN.Enabled = False
        '***
        Call ChargementVisualisation()
        '***
    End Sub

    Private Sub ArriveeVerification()
        Dim I As Integer
        If Not IsNumeric(SCH) Or SCH < 1 Or SCH > 20 Then
            MsgBox("donnée obligatoire et numérique" + vbCrLf + "comprise entre 1 et 20.", 48, "Erreur")
            Ok = False
        Else
            While I < (Pos - 1) And Pos > 1
                If SCH = ARQT(I) Then
                    Ok = False
                    MsgBox("Les données doivent être strictement différentes.", 48, "Erreur")
                    I = Pos - 1
                End If
                I += 1
            End While
        End If
    End Sub

    '***
    '*** VISUALISATION
    '***

    Private Sub PAGEPRE_Click(sender As Object, e As EventArgs) Handles PAGEPRE.Click
        Dim I As Integer
        '*** affichage des combinaisons quinte a l'ecran
        Call AfficheTous()
        LBox = (CPage - 1) * 70
        IBox = LBox - 70
        I = 0
        Do
            If BXT(0, IBox) > 0 Then
                Quintes(I).Text = CStr(BXT(0, IBox)) & "." & CStr(BXT(1, IBox)) & "." & CStr(BXT(2, IBox)) & "." & CStr(BXT(3, IBox)) & "." & CStr(BXT(4, IBox))
                Quintes(I).Visible = True
                CBox = IBox
            Else
                Quintes(I).Visible = False
            End If
            IBox += 1
            I += 1
        Loop Until IBox = LBox Or I = 70
        PAGESUI.Enabled = True
        CPage -= 1
        NBPAGE.Text = CPage & "/" & TPage
        If CPage = 1 Then
            PAGEPRE.Enabled = False
        End If
        If CPage = TPage And TPage = 1 Then
            PAGEPRE.Enabled = False
            PAGESUI.Enabled = False
        End If
    End Sub

    Private Sub PAGESUI_Click(sender As Object, e As EventArgs) Handles PAGESUI.Click
        Dim I As Integer
        Call AfficheTous()
        '*** Dernier Index Quinte Affiché : CBox
        LBox = (CPage + 1) * 70
        IBox = LBox - 70
        I = 0
        Do
            If BXT(0, IBox) > 0 Then
                Quintes(I).Text = CStr(BXT(0, IBox)) & "." & CStr(BXT(1, IBox)) & "." & CStr(BXT(2, IBox)) & "." & CStr(BXT(3, IBox)) & "." & CStr(BXT(4, IBox))
                Quintes(I).Visible = True
                CBox = IBox
            Else
                Quintes(I).Text = ""
                Quintes(I).Visible = False
            End If
            IBox += 1
            I += 1
        Loop Until IBox = LBox Or I = 70
        CPage += 1
        NBPAGE.Text = CPage & "/" & TPage
        If CPage = TPage And TPage = 1 Then
            PAGEPRE.Enabled = False
            PAGESUI.Enabled = False
        ElseIf CPage = TPage Then
            PAGEPRE.Enabled = True
            PAGESUI.Enabled = False
        ElseIf CPage > 1 Then
            PAGEPRE.Enabled = True
        End If
    End Sub

    Private Sub ChargementVisualisation()
        Dim I As Integer
        Dim CptCombi As Integer
        Dim CptPages As Single
        IBox = 0
        '*** Compte le nombre de combinaisons
        For I = 0 To 999
            If BXT(0, I) > 0 Then
                CptCombi += 1
            End If
        Next
        CptPages = CptCombi / 70
        '*** affichage du nombre de pages
        If TPage <> CptPages Then
            TPage = Math.Truncate(CptCombi / 70) + 1
        End If
        '*** affichage des combinaisons quinte a l'ecran
        Call AfficheTous()
        I = 0
        Do
            If BXT(0, IBox) > 0 Then
                Quintes(I).Text = CStr(BXT(0, IBox)) & "." & CStr(BXT(1, IBox)) & "." & CStr(BXT(2, IBox)) & "." & CStr(BXT(3, IBox)) & "." & CStr(BXT(4, IBox))
                '*** Dernier Index Quinte Affiché : CBox
                CBox = IBox
            Else
                Quintes(I).Visible = False
            End If
            IBox += 1
            I += 1
        Loop Until IBox = 70
        PAGEPRE.Enabled = False
        If CBox < 69 Then
            PAGESUI.Enabled = False
        Else
            PAGESUI.Enabled = True
        End If
        CPage = 1
        NBPAGE.Text = CPage & "/" & TPage
    End Sub

    Private Sub IMPRESSION_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ZeTurf_Click(sender As Object, e As EventArgs)

    End Sub

    Sub AfficheTous()
        Dim I As Integer
        For I = 0 To 69
            Quintes(I).Visible = True
        Next I
    End Sub

    '***
    '*** Outils et Fonctions
    '***
    '*** Tri pour optimisation des champs réduits
    '***

    Private Sub OptimisationChampsReduit(RD1 As Integer, RD2 As Integer, RD3 As Integer, RD4 As Integer, RD5 As Integer)
        Dim CCH(4, 13) As Long
        Dim REP(4) As Long
        Dim REF As Long
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer
        '*** Comptabilisation du nombre de repetitions des 14
        '*** chevaux dans chaque position
        For I = 0 To 999
            For J = 0 To 4
                For K = 0 To 13
                    If TRI(J, I) = B(K) Then
                        CCH(J, K) = CCH(J, K) + 1
                        K = 13
                    End If
                Next K
            Next J
            If TRI(0, I) < 1 Then
                I = 999
            End If
        Next I

        '*** Comptabilisation du nombre de chevaux les plus cites
        '*** pour chaque position
        For I = 0 To 4
            For J = 0 To 13
                REP(I) = REP(I) + (CCH(I, J) * CCH(I, J))
            Next J
        Next I

        '*** Recherche de la position la plus favorable
        '*** pour le champs reduit
        REF = REP(0)
        RD5 = 1
        For I = 2 To 5
            If REF >= REP(I - 1) Then
                REF = REP(I - 1)
                RD5 = I
            End If
        Next I

        '*** Recherche des criteres de tri les plus favorables
        '*** pour l'optimisation du nombre de champs reduits
        REP(RD5 - 1) = 0
        REF = REP(0)
        RD1 = 1
        For I = 2 To 5
            If REF <= REP(I - 1) Then
                REF = REP(I - 1)
                RD1 = I
            End If
        Next I
        REP(RD1 - 1) = 0
        REF = REP(0)
        RD2 = 1
        For I = 2 To 5
            If REF <= REP(I - 1) Then
                REF = REP(I - 1)
                RD2 = I
            End If
        Next I
        REP(RD2 - 1) = 0
        REF = REP(0)
        RD3 = 1
        For I = 2 To 5
            If REF <= REP(I - 1) Then
                REF = REP(I - 1)
                RD3 = I
            End If
        Next I
        REP(RD3 - 1) = 0
        REF = REP(0)
        RD4 = 1
        For I = 2 To 5
            If REF <= REP(I - 1) Then
                REF = REP(I - 1)
                RD4 = I
            End If
        Next I
    End Sub

End Class