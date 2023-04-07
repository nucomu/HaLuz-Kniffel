Option Explicit On
'Option Strict On

Public Class frmKniffel
    '----------VARIABLENDEKLARATIONEN---------------
    'Schleifenzähler
    Dim i, j As Short

    'Punktesummen
    Dim shEndsumme, shPunkteOben1, shPunkteOben2, shPunkteUnten, shBonus As Short
    Dim booBonus As Boolean = False

    'Einzelpunkte
    Dim shEinser, shZweier, shDreier, shVierer, shFünfer, shSechser As Short
    Dim shDreipasch, shVierpasch, shFullhouse, shKlstr, shGrstr, shKniffel, shChance As Short

    'Bereits belegte Felder
    Dim booCheckEinser, booCheckZweier, booCheckDreier, booCheckVierer, booCheckFünfer, booCheckSechser As Boolean
    Dim booCheckDreipasch, booCheckVierpasch, booCheckFullhouse, booCheckKlstr, booCheckGrstr, booCheckKniffel, booCheckChance As Boolean

    'Anzahl Würfe
    Dim shWürfe As Short

    'Würfel
    Dim W1, W2, W3, W4, W5 As Short                     'Würfelwerte
    Dim booW1, booW2, booW3, booW4, booW5 As Boolean    'gesicherte Würfel

    'Spielende
    Dim booEnde As Boolean = False

    'Protokollierung
    Dim strProtokoll As String
    Dim strProtokollFilename As String

    '----------EINSTELLUNGEN---------------
    'Konstanten
    Const shMaxWürfe As Short = 3                       'Maximal Anzahl Würfe
    Const shPktBonus As Short = 35                      'Bonus
    Const shSchwelle As Short = 63                      'Bonusschwelle
    Const shPktKlStr As Short = 30
    Const shPktgrStr As Short = 40
    Const shPktFullHouse As Short = 25
    Const shPktKniffel As Short = 50

    Const shZähldauer As Short = 25

    Private Sub txtProtokoll_TextChanged(sender As Object, e As EventArgs) Handles txtProtokoll.TextChanged
        txtProtokoll.SelectionStart = txtProtokoll.TextLength  ' Cursor an das Ende
        txtProtokoll.ScrollToCaret()                           ' Anzeige des Testes an der Cursorposition
    End Sub

    Private Sub Form1_load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Initialisierung
        Randomize()                     'Zufallsgenerator an

        'Endsumme nullen
        lblPunkteEndsumme.Text = "0"

        strProtokoll = "Spiel am/um " & DateTime.Now.ToString & vbCrLf & vbCrLf &
                       "Maximale Wurfanzahl:      " & shMaxWürfe.ToString & vbCrLf &
                       "Bonuspunkte:              " & shPktBonus.ToString & vbCrLf &
                       "Schwelle für Bonuspunkte: " & shSchwelle.ToString & vbCrLf &
                       "Kleine Straße:            " & shPktKlStr.ToString & vbCrLf &
                       "Große Straße:             " & shPktgrStr.ToString & vbCrLf &
                       "Full House:               " & shPktFullHouse.ToString & vbCrLf &
                       "Kniffel:                  " & shPktKniffel.ToString & vbCrLf
        txtProtokoll.Text = strProtokoll

        'OK-Buttons erst mal ausgegraut
        btnOkEinser.Enabled = False
        btnOkZweier.Enabled = False
        btnOkDreier.Enabled = False
        btnOkVierer.Enabled = False
        btnOkFünfer.Enabled = False
        btnOkSechser.Enabled = False
        btnOkDreipasch.Enabled = False
        btnOkVierpasch.Enabled = False
        btnOkFullhouse.Enabled = False
        btnOkKlstr.Enabled = False
        btnOkGrstr.Enabled = False
        btnOkKniffel.Enabled = False
        btnOkChance.Enabled = False

    End Sub

    'Würfelsteuerung
    Private Sub btnW1_Click(sender As Object, e As EventArgs) Handles btnW1.Click
        'Würfel 1
        If Not booW1 Then
            Select Case W1 'Wert bestimmen, Grüner Würfel anzeigen
                Case 1
                    btnW1.BackgroundImage = My.Resources._1g
                Case 2
                    btnW1.BackgroundImage = My.Resources._2g
                Case 3
                    btnW1.BackgroundImage = My.Resources._3g
                Case 4
                    btnW1.BackgroundImage = My.Resources._4g
                Case 5
                    btnW1.BackgroundImage = My.Resources._5g
                Case 6
                    btnW1.BackgroundImage = My.Resources._6g
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 1 mit dem Wert " & W1 & " ausgewählt."
            txtProtokoll.Text = strProtokoll
            booW1 = True
        Else
            Select Case W1 'Wert bestimmen, Weißer Würfel anzeigen
                Case 1
                    btnW1.BackgroundImage = My.Resources._1
                Case 2
                    btnW1.BackgroundImage = My.Resources._2
                Case 3
                    btnW1.BackgroundImage = My.Resources._3
                Case 4
                    btnW1.BackgroundImage = My.Resources._4
                Case 5
                    btnW1.BackgroundImage = My.Resources._5
                Case 6
                    btnW1.BackgroundImage = My.Resources._6
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 1 mit dem Wert " & W1 & " abgewählt."
            txtProtokoll.Text = strProtokoll
            booW1 = False
        End If
    End Sub

    Private Sub btnW2_Click(sender As Object, e As EventArgs) Handles btnW2.Click
        'Würfel 2
        If Not booW2 Then
            Select Case W2 'Wert bestimmen, Grüner Würfel anzeigen
                Case 1
                    btnW2.BackgroundImage = My.Resources._1g
                Case 2
                    btnW2.BackgroundImage = My.Resources._2g
                Case 3
                    btnW2.BackgroundImage = My.Resources._3g
                Case 4
                    btnW2.BackgroundImage = My.Resources._4g
                Case 5
                    btnW2.BackgroundImage = My.Resources._5g
                Case 6
                    btnW2.BackgroundImage = My.Resources._6g
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 2 mit dem Wert " & W2 & " ausgewählt."
            txtProtokoll.Text = strProtokoll
            booW2 = True
        Else
            Select Case W2 'Wert bestimmen, Weißer Würfel anzeigen
                Case 1
                    btnW2.BackgroundImage = My.Resources._1
                Case 2
                    btnW2.BackgroundImage = My.Resources._2
                Case 3
                    btnW2.BackgroundImage = My.Resources._3
                Case 4
                    btnW2.BackgroundImage = My.Resources._4
                Case 5
                    btnW2.BackgroundImage = My.Resources._5
                Case 6
                    btnW2.BackgroundImage = My.Resources._6
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 2 mit dem Wert " & W2 & " abgewählt."
            txtProtokoll.Text = strProtokoll
            booW2 = False
        End If
    End Sub

    Private Sub btnW3_Click(sender As Object, e As EventArgs) Handles btnW3.Click
        'Würfel 3
        If Not booW3 Then
            Select Case W3 'Wert bestimmen, Grüner Würfel anzeigen
                Case 1
                    btnW3.BackgroundImage = My.Resources._1g
                Case 2
                    btnW3.BackgroundImage = My.Resources._2g
                Case 3
                    btnW3.BackgroundImage = My.Resources._3g
                Case 4
                    btnW3.BackgroundImage = My.Resources._4g
                Case 5
                    btnW3.BackgroundImage = My.Resources._5g
                Case 6
                    btnW3.BackgroundImage = My.Resources._6g
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 3 mit dem Wert " & W3 & " ausgewählt."
            txtProtokoll.Text = strProtokoll
            booW3 = True
        Else
            Select Case W3 'Wert bestimmen, Weißer Würfel anzeigen
                Case 1
                    btnW3.BackgroundImage = My.Resources._1
                Case 2
                    btnW3.BackgroundImage = My.Resources._2
                Case 3
                    btnW3.BackgroundImage = My.Resources._3
                Case 4
                    btnW3.BackgroundImage = My.Resources._4
                Case 5
                    btnW3.BackgroundImage = My.Resources._5
                Case 6
                    btnW3.BackgroundImage = My.Resources._6
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 3 mit dem Wert " & W3 & " abgewählt."
            txtProtokoll.Text = strProtokoll
            booW3 = False
        End If
    End Sub

    Private Sub btnW4_Click(sender As Object, e As EventArgs) Handles btnW4.Click
        'Würfel 4
        If Not booW4 Then
            Select Case W4 'Wert bestimmen, Grüner Würfel anzeigen
                Case 1
                    btnW4.BackgroundImage = My.Resources._1g
                Case 2
                    btnW4.BackgroundImage = My.Resources._2g
                Case 3
                    btnW4.BackgroundImage = My.Resources._3g
                Case 4
                    btnW4.BackgroundImage = My.Resources._4g
                Case 5
                    btnW4.BackgroundImage = My.Resources._5g
                Case 6
                    btnW4.BackgroundImage = My.Resources._6g
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 4 mit dem Wert " & W4 & " ausgewählt."
            txtProtokoll.Text = strProtokoll
            booW4 = True
        Else
            Select Case W4 'Wert bestimmen, Weißer Würfel anzeigen
                Case 1
                    btnW4.BackgroundImage = My.Resources._1
                Case 2
                    btnW4.BackgroundImage = My.Resources._2
                Case 3
                    btnW4.BackgroundImage = My.Resources._3
                Case 4
                    btnW4.BackgroundImage = My.Resources._4
                Case 5
                    btnW4.BackgroundImage = My.Resources._5
                Case 6
                    btnW4.BackgroundImage = My.Resources._6
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 4 mit dem Wert " & W4 & " abgewählt."
            txtProtokoll.Text = strProtokoll
            booW4 = False
        End If
    End Sub

    Private Sub btnW5_Click(sender As Object, e As EventArgs) Handles btnW5.Click
        'Würfel 5
        If Not booW5 Then
            Select Case W5 'Wert bestimmen, Grüner Würfel anzeigen
                Case 1
                    btnW5.BackgroundImage = My.Resources._1g
                Case 2
                    btnW5.BackgroundImage = My.Resources._2g
                Case 3
                    btnW5.BackgroundImage = My.Resources._3g
                Case 4
                    btnW5.BackgroundImage = My.Resources._4g
                Case 5
                    btnW5.BackgroundImage = My.Resources._5g
                Case 6
                    btnW5.BackgroundImage = My.Resources._6g
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 5 mit dem Wert " & W5 & " ausgewählt."
            txtProtokoll.Text = strProtokoll
            booW5 = True
        Else
            Select Case W5 'Wert bestimmen, Weißer Würfel anzeigen
                Case 1
                    btnW5.BackgroundImage = My.Resources._1
                Case 2
                    btnW5.BackgroundImage = My.Resources._2
                Case 3
                    btnW5.BackgroundImage = My.Resources._3
                Case 4
                    btnW5.BackgroundImage = My.Resources._4
                Case 5
                    btnW5.BackgroundImage = My.Resources._5
                Case 6
                    btnW5.BackgroundImage = My.Resources._6
            End Select
            strProtokoll = strProtokoll & vbCrLf & "Würfel 5 mit dem Wert " & W5 & " abgewählt."
            txtProtokoll.Text = strProtokoll
            booW5 = False
        End If
    End Sub

    'Punkteauswertung:
    Private Sub btnOkEinser_Click(sender As Object, e As EventArgs) Handles btnOkEinser.Click
        shEinser = CShort(lblPunkteVorschauEinser.Text)     'Vorschau in Variable übernehmen
        lblPunkteEinser.Text = shEinser.ToString            'Punktzahl werten
        btnOkEinser.Visible = False                         'benutzten Button ausblenden
        lblPunkteVorschauEinser.Visible = False             'Vorschaufeld ausblenden
        booCheckEinser = True                               'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "EINSER ausgewählt, " & shEinser.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkZweier_Click(sender As Object, e As EventArgs) Handles btnOkZweier.Click
        shZweier = CShort(lblPunkteVorschauZweier.Text)     'Vorschau in Variable übernehmen
        lblPunkteZweier.Text = shZweier.ToString            'Punktzahl werten
        btnOkZweier.Visible = False                         'benutzten Button ausblenden
        lblPunkteVorschauZweier.Visible = False             'Vorschaufeld ausblenden
        booCheckZweier = True                               'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "ZWEIER ausgewählt, " & shZweier.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkDreier_Click(sender As Object, e As EventArgs) Handles btnOkDreier.Click
        shDreier = CShort(lblPunkteVorschauDreier.Text)     'Vorschau in Variable übernehmen
        lblPunkteDreier.Text = shDreier.ToString            'Punktzahl werten
        btnOkDreier.Visible = False                         'benutzten Button ausblenden
        lblPunkteVorschauDreier.Visible = False             'Vorschaufeld ausblenden
        booCheckDreier = True                               'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "DREIER ausgewählt, " & shDreier.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkVierer_Click(sender As Object, e As EventArgs) Handles btnOkVierer.Click
        shVierer = CShort(lblPunkteVorschauVierer.Text)     'Vorschau in Variable übernehmen
        lblPunkteVierer.Text = shVierer.ToString            'Punktzahl werten
        btnOkVierer.Visible = False                         'benutzten Button ausblenden
        lblPunkteVorschauVierer.Visible = False             'Vorschaufeld ausblenden
        booCheckVierer = True                               'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "VIERER ausgewählt, " & shVierer.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkFünfer_Click(sender As Object, e As EventArgs) Handles btnOkFünfer.Click
        shFünfer = CShort(lblPunkteVorschauFünfer.Text)     'Vorschau in Variable übernehmen
        lblPunkteFünfer.Text = shFünfer.ToString            'Punktzahl werten
        btnOkFünfer.Visible = False                         'benutzten Button ausblenden
        lblPunkteVorschauFünfer.Visible = False             'Vorschaufeld ausblenden
        booCheckFünfer = True                               'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "FÜNFER ausgewählt, " & shFünfer.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkSechser_Click(sender As Object, e As EventArgs) Handles btnOkSechser.Click
        shSechser = CShort(lblPunkteVorschauSechser.Text)   'Vorschau in Variable übernehmen
        lblPunkteSechser.Text = shSechser.ToString          'Punktzahl werten
        btnOkSechser.Visible = False                        'benutzten Button ausblenden
        lblPunkteVorschauSechser.Visible = False            'Vorschaufeld ausblenden
        booCheckSechser = True                              'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "SECHSER ausgewählt, " & shSechser.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkDreipasch_Click(sender As Object, e As EventArgs) Handles btnOkDreipasch.Click
        shDreipasch = CShort(lblPunkteVorschauDreipasch.Text) 'Vorschau in Variable übernehmen
        lblPunkteDreipasch.Text = shDreipasch.ToString        'Punktzahl werten
        btnOkDreipasch.Visible = False                      'benutzten Button ausblenden
        lblPunkteVorschauDreipasch.Visible = False          'Vorschaufeld ausblenden
        booCheckDreipasch = True                            'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "DREIERPASCH ausgewählt, " & shDreipasch.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkVierpasch_Click(sender As Object, e As EventArgs) Handles btnOkVierpasch.Click
        shVierpasch = CShort(lblPunkteVorschauVierpasch.Text) 'Vorschau in Variable übernehmen
        lblPunkteVierpasch.Text = shVierpasch.ToString        'Punktzahl werten
        btnOkVierpasch.Visible = False                      'benutzten Button ausblenden
        lblPunkteVorschauVierpasch.Visible = False          'Vorschaufeld ausblenden
        booCheckVierpasch = True                            'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "VIERERPASCH ausgewählt, " & shVierpasch.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkFullhouse_Click(sender As Object, e As EventArgs) Handles btnOkFullhouse.Click
        shFullhouse = CShort(lblPunkteVorschauFullhouse.Text) 'Vorschau in Variable übernehmen
        lblPunkteFullhouse.Text = shFullhouse.ToString        'Punktzahl werten
        btnOkFullhouse.Visible = False                      'benutzten Button ausblenden
        lblPunkteVorschauFullhouse.Visible = False          'Vorschaufeld ausblenden
        booCheckFullhouse = True                            'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "FULLHOUSE ausgewählt, " & shFullhouse.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkKlstr_Click(sender As Object, e As EventArgs) Handles btnOkKlstr.Click
        shKlstr = CShort(lblPunkteVorschauKlstr.Text)       'Vorschau in Variable übernehmen
        lblPunkteKlstr.Text = shKlstr.ToString              'Punktzahl werten
        btnOkKlstr.Visible = False                          'benutzten Button ausblenden
        lblPunkteVorschauKlstr.Visible = False              'Vorschaufeld ausblenden
        booCheckKlstr = True                                'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "KLEINE STRAßE ausgewählt, " & shKlstr.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkGrstr_Click(sender As Object, e As EventArgs) Handles btnOkGrstr.Click
        shGrstr = CShort(lblPunkteVorschauGrstr.Text)       'Vorschau in Variable übernehmen
        lblPunkteGrstr.Text = shGrstr.ToString              'Punktzahl werten
        btnOkGrstr.Visible = False                          'benutzten Button ausblenden
        lblPunkteVorschauGrstr.Visible = False              'Vorschaufeld ausblenden
        booCheckGrstr = True                                'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "GROßE STRAßE ausgewählt, " & shGrstr.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkKniffel_Click(sender As Object, e As EventArgs) Handles btnOkKniffel.Click
        shKniffel = CShort(lblPunkteVorschauKniffel.Text)   'Vorschau in Variable übernehmen
        lblPunkteKniffel.Text = shKniffel.ToString          'Punktzahl werten
        btnOkKniffel.Visible = False                        'benutzten Button ausblenden
        lblPunkteVorschauKniffel.Visible = False            'Vorschaufeld ausblenden
        booCheckKniffel = True                              'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "KNIFFEL ausgewählt, " & shKniffel.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Sub btnOkChance_Click(sender As Object, e As EventArgs) Handles btnOkChance.Click
        shChance = CShort(lblPunkteVorschauChance.Text)     'Vorschau in Variable übernehmen
        lblPunkteChance.Text = shChance.ToString            'Punktzahl werten
        btnOkChance.Visible = False                         'benutzten Button ausblenden
        lblPunkteVorschauChance.Visible = False             'Vorschaufeld ausblenden
        booCheckChance = True                               'Feld gespielt
        strProtokoll = strProtokoll & vbCrLf & "CHANCE ausgewählt, " & shChance.ToString & " Punkte gutgeschrieben."
        txtProtokoll.Text = strProtokoll
        AuswahlSperren()                                    'alle anderen blockieren
        Berechnen()                                         'Gesamtpunktzahl berechnen
        Würfelfreigabe()                                    'Würfel für nächsten Zug freigeben
    End Sub

    Private Async Sub btnWürfeln_Click(sender As Object, e As EventArgs) Handles btnWürfeln.Click
        'Neustart bei Spielende
        If booEnde Then
            booEnde = False
            Neustart()
        End If

        'beim ersten Wurf alle Würfel freigeben
        If shWürfe = 0 Then
            booW1 = False
            booW2 = False
            booW3 = False
            booW4 = False
            booW5 = False
        End If

        'Würfe hochzählen
        shWürfe += CShort(1)

        'Würfel während des Wurfes ausgrauen und deaktivieren
        btnWürfeln.Enabled = False

        btnW1.Enabled = False
        btnW2.Enabled = False
        btnW3.Enabled = False
        btnW4.Enabled = False
        btnW5.Enabled = False

        'Würfeln inkl. Animation
        For i = 1 To shZähldauer
            Würfeln()
            Await Task.Delay(shZähldauer)
        Next i

        strProtokoll = strProtokoll & vbCrLf & shWürfe & ". WÜRFELN"
        strProtokoll = strProtokoll & vbCrLf & W1.ToString & ", " &
                                               W2.ToString & ", " &
                                               W3.ToString & ", " &
                                               W4.ToString & ", " &
                                               W5.ToString
        txtProtokoll.Text = strProtokoll

        'freie Würfel wieder weiß machen
        If Not booW1 Then
            Select Case W1
                Case 1
                    btnW1.BackgroundImage = My.Resources._1
                Case 2
                    btnW1.BackgroundImage = My.Resources._2
                Case 3
                    btnW1.BackgroundImage = My.Resources._3
                Case 4
                    btnW1.BackgroundImage = My.Resources._4
                Case 5
                    btnW1.BackgroundImage = My.Resources._5
                Case 6
                    btnW1.BackgroundImage = My.Resources._6
            End Select
        End If

        If Not booW2 Then
            Select Case W2
                Case 1
                    btnW2.BackgroundImage = My.Resources._1
                Case 2
                    btnW2.BackgroundImage = My.Resources._2
                Case 3
                    btnW2.BackgroundImage = My.Resources._3
                Case 4
                    btnW2.BackgroundImage = My.Resources._4
                Case 5
                    btnW2.BackgroundImage = My.Resources._5
                Case 6
                    btnW2.BackgroundImage = My.Resources._6
            End Select
        End If

        If Not booW3 Then
            Select Case W3
                Case 1
                    btnW3.BackgroundImage = My.Resources._1
                Case 2
                    btnW3.BackgroundImage = My.Resources._2
                Case 3
                    btnW3.BackgroundImage = My.Resources._3
                Case 4
                    btnW3.BackgroundImage = My.Resources._4
                Case 5
                    btnW3.BackgroundImage = My.Resources._5
                Case 6
                    btnW3.BackgroundImage = My.Resources._6
            End Select
        End If

        If Not booW4 Then
            Select Case W4
                Case 1
                    btnW4.BackgroundImage = My.Resources._1
                Case 2
                    btnW4.BackgroundImage = My.Resources._2
                Case 3
                    btnW4.BackgroundImage = My.Resources._3
                Case 4
                    btnW4.BackgroundImage = My.Resources._4
                Case 5
                    btnW4.BackgroundImage = My.Resources._5
                Case 6
                    btnW4.BackgroundImage = My.Resources._6
            End Select
        End If

        If Not booW5 Then
            Select Case W5
                Case 1
                    btnW5.BackgroundImage = My.Resources._1
                Case 2
                    btnW5.BackgroundImage = My.Resources._2
                Case 3
                    btnW5.BackgroundImage = My.Resources._3
                Case 4
                    btnW5.BackgroundImage = My.Resources._4
                Case 5
                    btnW5.BackgroundImage = My.Resources._5
                Case 6
                    btnW5.BackgroundImage = My.Resources._6
            End Select
        End If

        'Würfel wieder freigeben
        btnW1.Enabled = True
        btnW2.Enabled = True
        btnW3.Enabled = True
        btnW4.Enabled = True
        btnW5.Enabled = True
        btnWürfeln.Enabled = True

        'Freigabe der unbenutzen OK-Buttons
        If Not booCheckEinser Then btnOkEinser.Enabled = True
        If Not booCheckZweier Then btnOkZweier.Enabled = True
        If Not booCheckDreier Then btnOkDreier.Enabled = True
        If Not booCheckVierer Then btnOkVierer.Enabled = True
        If Not booCheckFünfer Then btnOkFünfer.Enabled = True
        If Not booCheckSechser Then btnOkSechser.Enabled = True
        If Not booCheckDreipasch Then btnOkDreipasch.Enabled = True
        If Not booCheckVierpasch Then btnOkVierpasch.Enabled = True
        If Not booCheckFullhouse Then btnOkFullhouse.Enabled = True
        If Not booCheckKlstr Then btnOkKlstr.Enabled = True
        If Not booCheckGrstr Then btnOkGrstr.Enabled = True
        If Not booCheckKniffel Then btnOkKniffel.Enabled = True
        If Not booCheckChance Then btnOkChance.Enabled = True

        'Was wurde gewürfelt?
        Auswerten()

        'Würfelbutton: Beschriftung aktualisieren
        btnWürfeln.Text = "Würfeln! (" & shWürfe + 1 & "/" & shMaxWürfe & ")"

        'Würfebegrenzer
        If shWürfe >= shMaxWürfe Then
            btnWürfeln.Enabled = False
            btnWürfeln.Text = "Würfeln! (" & shWürfe & "+)"
        End If
    End Sub

    'Eigene Subs
    Async Sub Berechnen()       'Punkte berechnen
        'Punkte oberer Teil
        shPunkteOben1 = shEinser + shZweier + shDreier + shVierer + shFünfer + shSechser

        'Bonus?
        If Not booBonus Then
            If shPunkteOben1 >= shSchwelle Then
                shBonus = shPktBonus 'Else shBonus = 0
                lblPunkteBonus.Text = shBonus
                booBonus = True
                strProtokoll = strProtokoll & vbCrLf & shPunkteOben1.ToString & " Punkte, BONUS erreicht, " & shBonus.ToString & " Punkte gutgeschrieben."
                txtProtokoll.Text = strProtokoll
            End If
        End If

        'Gesamtpunkte oberer Teil
        shPunkteOben2 = shPunkteOben1 + shBonus

        'Gesamtpunkte unterer Teil
        shPunkteUnten = shDreipasch + shVierpasch + shFullhouse + shKlstr + shGrstr + shKniffel + shChance

        'Endsumme
        shEndsumme = shPunkteOben2 + shPunkteUnten

        'Punkte anzeigen
        lblPunkteEinser.Text = shEinser
        lblPunkteZweier.Text = shZweier
        lblPunkteDreier.Text = shDreier
        lblPunkteVierer.Text = shVierer
        lblPunkteFünfer.Text = shFünfer
        lblPunkteSechser.Text = shSechser

        lblPunkteDreipasch.Text = shDreipasch
        lblPunkteVierpasch.Text = shVierpasch
        lblPunkteFullhouse.Text = shFullhouse
        lblPunkteKlstr.Text = shKlstr
        lblPunkteGrstr.Text = shGrstr
        lblPunkteKniffel.Text = shKniffel
        lblPunkteChance.Text = shChance

        lblPunkteOben1.Text = shPunkteOben1
        lblPunkteOben2.Text = shPunkteOben2
        lblPunkteUnten.Text = shPunkteUnten

        strProtokoll = strProtokoll & vbCrLf & "GESAMTSUMME beträgt " & shEndsumme.ToString & vbCrLf
        txtProtokoll.Text = strProtokoll

        'Punktezähler (animiert)
        If lblPunkteEndsumme.Text < shEndsumme Then
            For i = lblPunkteEndsumme.Text + 1 To shEndsumme
                lblPunkteEndsumme.Text += 1
                Await Task.Delay(shZähldauer)
                'bei zu schnellem Würfeln werden Punkte mehrfach berechnet. Das wird hier korrigiert: (würgaround)
                If lblPunkteEndsumme.Text = shEndsumme Then
                    lblPunkteEndsumme.Text = shEndsumme
                    i = shEndsumme
                End If
            Next
        End If

    End Sub

    Sub Würfeln()
        'Würfeln mit nicht-gesperrten Würfeln
        If Not booW1 Then W1 = CInt(Math.Floor((6) * Rnd())) + 1
        If Not booW2 Then W2 = CInt(Math.Floor((6) * Rnd())) + 1
        If Not booW3 Then W3 = CInt(Math.Floor((6) * Rnd())) + 1
        If Not booW4 Then W4 = CInt(Math.Floor((6) * Rnd())) + 1
        If Not booW5 Then W5 = CInt(Math.Floor((6) * Rnd())) + 1

        If Not booW1 Then
            Select Case W1
                Case 1
                    btnW1.BackgroundImage = My.Resources._1ge
                Case 2
                    btnW1.BackgroundImage = My.Resources._2ge
                Case 3
                    btnW1.BackgroundImage = My.Resources._3ge
                Case 4
                    btnW1.BackgroundImage = My.Resources._4ge
                Case 5
                    btnW1.BackgroundImage = My.Resources._5ge
                Case 6
                    btnW1.BackgroundImage = My.Resources._6ge
            End Select
        End If

        If Not booW2 Then
            Select Case W2
                Case 1
                    btnW2.BackgroundImage = My.Resources._1ge
                Case 2
                    btnW2.BackgroundImage = My.Resources._2ge
                Case 3
                    btnW2.BackgroundImage = My.Resources._3ge
                Case 4
                    btnW2.BackgroundImage = My.Resources._4ge
                Case 5
                    btnW2.BackgroundImage = My.Resources._5ge
                Case 6
                    btnW2.BackgroundImage = My.Resources._6ge
            End Select
        End If

        If Not booW3 Then
            Select Case W3
                Case 1
                    btnW3.BackgroundImage = My.Resources._1ge
                Case 2
                    btnW3.BackgroundImage = My.Resources._2ge
                Case 3
                    btnW3.BackgroundImage = My.Resources._3ge
                Case 4
                    btnW3.BackgroundImage = My.Resources._4ge
                Case 5
                    btnW3.BackgroundImage = My.Resources._5ge
                Case 6
                    btnW3.BackgroundImage = My.Resources._6ge
            End Select
        End If

        If Not booW4 Then
            Select Case W4
                Case 1
                    btnW4.BackgroundImage = My.Resources._1ge
                Case 2
                    btnW4.BackgroundImage = My.Resources._2ge
                Case 3
                    btnW4.BackgroundImage = My.Resources._3ge
                Case 4
                    btnW4.BackgroundImage = My.Resources._4ge
                Case 5
                    btnW4.BackgroundImage = My.Resources._5ge
                Case 6
                    btnW4.BackgroundImage = My.Resources._6ge
            End Select
        End If

        If Not booW5 Then
            Select Case W5
                Case 1
                    btnW5.BackgroundImage = My.Resources._1ge
                Case 2
                    btnW5.BackgroundImage = My.Resources._2ge
                Case 3
                    btnW5.BackgroundImage = My.Resources._3ge
                Case 4
                    btnW5.BackgroundImage = My.Resources._4ge
                Case 5
                    btnW5.BackgroundImage = My.Resources._5ge
                Case 6
                    btnW5.BackgroundImage = My.Resources._6ge
            End Select
        End If
    End Sub

    Sub Auswerten() 'Was wurde gewürfelt?
        'Augen speichern
        Dim shWerte(5) As Short
        Dim shSortWerte(5) As Short
        shWerte(1) = W1
        shWerte(2) = W2
        shWerte(3) = W3
        shWerte(4) = W4
        shWerte(5) = W5

        shSortWerte = shWerte
        Sortieren(shSortWerte) 'Werte sortieren

        'Punktezähler für oberer Bereich
        Dim shOBeinser As Short = 0
        Dim shOBzweier As Short = 0
        Dim shOBdreier As Short = 0
        Dim shOBvierer As Short = 0
        Dim shOBfünfer As Short = 0
        Dim shOBsechser As Short = 0

        For i = 1 To 5
            If shSortWerte(i) = 1 Then shOBeinser += 1
            If shSortWerte(i) = 2 Then shOBzweier += 2
            If shSortWerte(i) = 3 Then shOBdreier += 3
            If shSortWerte(i) = 4 Then shOBvierer += 4
            If shSortWerte(i) = 5 Then shOBfünfer += 5
            If shSortWerte(i) = 6 Then shOBsechser += 6
        Next

        lblPunkteVorschauEinser.Text = shOBeinser
        lblPunkteVorschauZweier.Text = shOBzweier
        lblPunkteVorschauDreier.Text = shOBdreier
        lblPunkteVorschauVierer.Text = shOBvierer
        lblPunkteVorschauFünfer.Text = shOBfünfer
        lblPunkteVorschauSechser.Text = shOBsechser

        'Unterer Bereich

        'Chance?
        lblPunkteVorschauChance.Text = shSortWerte(1) + shSortWerte(2) + shSortWerte(3) + shSortWerte(4) + shSortWerte(5)

        'Kniffel?
        If shSortWerte(1) = shSortWerte(2) And shSortWerte(2) = shSortWerte(3) And shSortWerte(3) = shSortWerte(4) And shSortWerte(4) = shSortWerte(5) Then
            lblPunkteVorschauKniffel.Text = shPktKniffel
        Else
            lblPunkteVorschauKniffel.Text = 0
        End If

        'Große Straße?
        Dim booGrstr As Boolean = False
        'untere große Straße
        If shSortWerte(1) = 1 Then
            If shSortWerte(2) = 2 Then
                If shSortWerte(3) = 3 Then
                    If shSortWerte(4) = 4 Then
                        If shSortWerte(5) = 5 Then
                            booGrstr = True
                        End If
                    End If
                End If
            End If
        End If
        'obere große Straße
        If shSortWerte(1) = 2 Then
            If shSortWerte(2) = 3 Then
                If shSortWerte(3) = 4 Then
                    If shSortWerte(4) = 5 Then
                        If shSortWerte(5) = 6 Then
                            booGrstr = True
                        End If
                    End If
                End If
            End If
        End If
        If booGrstr Then lblPunkteVorschauGrstr.Text = shPktgrStr Else lblPunkteVorschauGrstr.Text = 0

        'Kleine Straße?
        Dim booKlstr As Boolean = False
        Dim booKlstrEins As Boolean = False
        Dim booKlstrZwei As Boolean = False
        Dim booKlstrDrei As Boolean = False
        Dim booKlstrVier As Boolean = False
        Dim booKlstrFünf As Boolean = False
        Dim booKlstrSechs As Boolean = False

        For i = 1 To 5
            If shSortWerte(i) = 1 Then booKlstrEins = True
            If shSortWerte(i) = 2 Then booKlstrZwei = True
            If shSortWerte(i) = 3 Then booKlstrDrei = True
            If shSortWerte(i) = 4 Then booKlstrVier = True
            If shSortWerte(i) = 5 Then booKlstrFünf = True
            If shSortWerte(i) = 6 Then booKlstrSechs = True
        Next

        If (booKlstrEins And booKlstrZwei And booKlstrDrei And booKlstrVier) Or
            (booKlstrZwei And booKlstrDrei And booKlstrVier And booKlstrFünf) Or
            (booKlstrDrei And booKlstrVier And booKlstrFünf And booKlstrSechs) Then
            booKlstr = True
        Else
            booKlstr = False
        End If

        If booKlstr Then lblPunkteVorschauKlstr.Text = shPktKlStr Else lblPunkteVorschauKlstr.Text = 0

        'Fullhouse
        If (shSortWerte(1) = shSortWerte(2) And shSortWerte(2) = shSortWerte(3) And shSortWerte(4) = shSortWerte(5)) Or
            (shSortWerte(1) = shSortWerte(2) And shSortWerte(3) = shSortWerte(4) And shSortWerte(4) = shSortWerte(5)) Then
            lblPunkteVorschauFullhouse.Text = shPktFullHouse
        Else
            lblPunkteVorschauFullhouse.Text = 0
        End If

        'Vierer Pasch
        If (shSortWerte(1) = shSortWerte(2) And shSortWerte(2) = shSortWerte(3) And shSortWerte(3) = shSortWerte(4)) Or
            (shSortWerte(2) = shSortWerte(3) And shSortWerte(3) = shSortWerte(4) And shSortWerte(4) = shSortWerte(5)) Then
            lblPunkteVorschauVierpasch.Text = shSortWerte(1) + shSortWerte(2) + shSortWerte(3) + shSortWerte(4) + shSortWerte(5)
        Else
            lblPunkteVorschauVierpasch.Text = 0
        End If

        'Dreier Pasch
        If (shSortWerte(1) = shSortWerte(2) And shSortWerte(2) = shSortWerte(3)) Or
            (shSortWerte(2) = shSortWerte(3) And shSortWerte(3) = shSortWerte(4)) Or
            (shSortWerte(3) = shSortWerte(4) And shSortWerte(4) = shSortWerte(5)) Then
            lblPunkteVorschauDreipasch.Text = shSortWerte(1) + shSortWerte(2) + shSortWerte(3) + shSortWerte(4) + shSortWerte(5)
        Else
            lblPunkteVorschauDreipasch.Text = 0
        End If

        'Berechnen
        Berechnen()
    End Sub

    Sub Würfelfreigabe()
        'nach OK alle Würfel freigeben
        booW1 = False
        booW2 = False
        booW3 = False
        booW4 = False
        booW5 = False

        btnWürfeln.Enabled = True
        btnWürfeln.Text = "Würfeln! (1" & "/" & shMaxWürfe & ")"
        shWürfe = 0

        'Spielende?
        If btnOkEinser.Visible = False And
            btnOkZweier.Visible = False And
            btnOkDreier.Visible = False And
            btnOkVierer.Visible = False And
            btnOkFünfer.Visible = False And
            btnOkSechser.Visible = False And
            btnOkDreipasch.Visible = False And
            btnOkVierpasch.Visible = False And
            btnOkFullhouse.Visible = False And
            btnOkKlstr.Visible = False And
            btnOkGrstr.Visible = False And
            btnOkKniffel.Visible = False And
            btnOkChance.Visible = False Then

            booEnde = True
            btnWürfeln.Text = "Spiel beendet! Klick für Neustart"

            strProtokoll = strProtokoll & vbCrLf & "Spiel beendet."
            strProtokoll = strProtokoll & vbCrLf & shEndsumme.ToString & " Punkte erreicht."
            txtProtokoll.Text = strProtokoll

            'Statistik speichern
            If shEndsumme > 250 Then
                My.Computer.FileSystem.WriteAllText("hiscore.txt",
                                                    shPunkteOben1 & vbTab & shBonus & vbTab & shPunkteOben2 & vbTab &
                                                    shPunkteUnten & vbTab & shEndsumme & vbTab &
                                                    Format(Now, "yyyy-MM-dd") & vbTab & Format(Now, "HH:mm:ss") & vbTab &
                                                    Environment.UserName & vbCrLf, append:=True)
            End If

            strProtokollFilename = DateAndTime.Year(Today).ToString & "-" &
                                   DateAndTime.Month(Today).ToString("00") & "-" &
                                   DateAndTime.Day(Today).ToString("00") & "_" &
                                   DateAndTime.Hour(Now).ToString("00") & "-" &
                                   DateAndTime.Minute(Now).ToString("00") & "-" &
                                   DateAndTime.Second(Now).ToString("00") & ".txt"
            My.Computer.FileSystem.WriteAllText(strProtokollFilename, strProtokoll, append:=False)
        End If
    End Sub

    Sub AuswahlSperren() 'alle OK ausblenden bis neu gewürfelt
        btnOkEinser.Enabled = False
        btnOkZweier.Enabled = False
        btnOkDreier.Enabled = False
        btnOkVierer.Enabled = False
        btnOkFünfer.Enabled = False
        btnOkSechser.Enabled = False
        btnOkDreipasch.Enabled = False
        btnOkVierpasch.Enabled = False
        btnOkFullhouse.Enabled = False
        btnOkKlstr.Enabled = False
        btnOkGrstr.Enabled = False
        btnOkKniffel.Enabled = False
        btnOkChance.Enabled = False
    End Sub

    Sub Neustart()
        'Variablen Rücksetzen
        i = 0
        j = 0

        'Punktesummen
        shEndsumme = 0
        lblPunkteEndsumme.Text = 0
        shPunkteOben1 = 0
        shPunkteOben2 = 0
        shPunkteUnten = 0
        shBonus = 0
        lblPunkteBonus.Text = 0
        booBonus = False

        'Einzelpunkte
        shEinser = 0
        shZweier = 0
        shDreier = 0
        shVierer = 0
        shFünfer = 0
        shSechser = 0
        shDreipasch = 0
        shVierpasch = 0
        shFullhouse = 0
        shKlstr = 0
        shGrstr = 0
        shKniffel = 0
        shChance = 0

        'Bereits belegte Felder
        booCheckEinser = False
        booCheckZweier = False
        booCheckDreier = False
        booCheckVierer = False
        booCheckFünfer = False
        booCheckSechser = False
        booCheckDreipasch = False
        booCheckVierpasch = False
        booCheckFullhouse = False
        booCheckKlstr = False
        booCheckGrstr = False
        booCheckKniffel = False
        booCheckChance = False

        'Anzahl Würfe
        shWürfe = 0

        'Würfel
        W1 = 0
        booW1 = False
        W2 = 0
        booW2 = False
        W3 = 0
        booW3 = False
        W4 = 0
        booW4 = False
        W5 = 0
        booW5 = False

        'Spielende
        booEnde = False

        'Buttons freigeben
        With btnOkEinser
            .Enabled = True
            .Visible = True
        End With
        With btnOkZweier
            .Enabled = True
            .Visible = True
        End With
        With btnOkDreier
            .Enabled = True
            .Visible = True
        End With
        With btnOkVierer
            .Enabled = True
            .Visible = True
        End With
        With btnOkFünfer
            .Enabled = True
            .Visible = True
        End With
        With btnOkSechser
            .Enabled = True
            .Visible = True
        End With
        With btnOkDreipasch
            .Enabled = True
            .Visible = True
        End With
        With btnOkVierpasch
            .Enabled = True
            .Visible = True
        End With
        With btnOkFullhouse
            .Enabled = True
            .Visible = True
        End With
        With btnOkKlstr
            .Enabled = True
            .Visible = True
        End With
        With btnOkGrstr
            .Enabled = True
            .Visible = True
        End With
        With btnOkKniffel
            .Enabled = True
            .Visible = True
        End With
        With btnOkChance
            .Enabled = True
            .Visible = True
        End With

        'Vorschau Labels rücksetzen
        With lblPunkteVorschauEinser
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauZweier
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauDreier
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauVierer
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauFünfer
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauSechser
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauDreipasch
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauVierpasch
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauFullhouse
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauKlstr
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauGrstr
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauKniffel
            .Text = 0
            .Visible = True
        End With
        With lblPunkteVorschauChance
            .Text = 0
            .Visible = True
        End With

        'Punkte Labels rücksetzen
        lblPunkteEinser.Text = 0
        lblPunkteZweier.Text = 0
        lblPunkteDreier.Text = 0
        lblPunkteVierer.Text = 0
        lblPunkteFünfer.Text = 0
        lblPunkteSechser.Text = 0
        lblPunkteDreipasch.Text = 0
        lblPunkteVierpasch.Text = 0
        lblPunkteFullhouse.Text = 0
        lblPunkteKlstr.Text = 0
        lblPunkteGrstr.Text = 0
        lblPunkteKniffel.Text = 0
        lblPunkteChance.Text = 0

        'OK-Buttons ausgrauen
        btnOkEinser.Enabled = False
        btnOkZweier.Enabled = False
        btnOkDreier.Enabled = False
        btnOkVierer.Enabled = False
        btnOkFünfer.Enabled = False
        btnOkSechser.Enabled = False
        btnOkDreipasch.Enabled = False
        btnOkVierpasch.Enabled = False
        btnOkFullhouse.Enabled = False
        btnOkKlstr.Enabled = False
        btnOkGrstr.Enabled = False
        btnOkKniffel.Enabled = False
        btnOkChance.Enabled = False

        'Protokoll löschen
        strProtokoll = ""
        txtProtokoll.Text = strProtokoll
    End Sub

    Sub Sortieren(array) 'Bubble Sort
        Dim vTemp As Short

        For j = UBound(array) - 1 To LBound(array) Step -1
            ' Alle links davon liegenden Zeichen auf richtige Sortierung der jeweiligen Nachfolger überprüfen: 
            For i = LBound(array) To j
                ' Ist das aktuelle Element seinem Nachfolger gegenüber korrekt sortiert? 
                If array(i) > array(i + 1) Then
                    ' Element und seinen Nachfolger vertauschen. 
                    vTemp = array(i)
                    array(i) = array(i + 1)
                    array(i + 1) = vTemp
                End If
            Next i
        Next j
    End Sub
End Class
