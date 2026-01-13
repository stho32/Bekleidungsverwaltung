Attribute VB_Name = "modBerechnung"
Option Explicit

' ============================================================================
' modBerechnung - Berechnungslogik für Ansprüche
' ============================================================================
' Gemäß App-Architecture/Excel-VBA-Conventions.md
' Enthält die Kernlogik für Anspruchsberechnung
' ============================================================================

' Konstante für Innendienst-Sonderregel
Private Const INNENDIENST_HEMD_DEFAULT As Integer = 2

' ----------------------------------------------------------------------------
' Effektiven Anspruch berechnen (inkl. Sonderregel Innendienst)
' ----------------------------------------------------------------------------
' Diese Funktion ermittelt den tatsächlichen Anspruch eines Mitarbeiters
' unter Berücksichtigung von:
' 1. Standard-Anspruch aus Sortiment
' 2. Sonderregel Innendienst (nur 2 Hemden/Blusen)
' 3. Individuelle Abweichungen (falls implementiert)
' ----------------------------------------------------------------------------
Public Function BerechneEffektivenAnspruch(lngPersonalnummer As Long, _
                                           intArtikelID As Integer) As Integer
    Dim intStandard As Integer
    Dim strBereich As String
    Dim strArtikelName As String
    Dim intInnendienstAnspruch As Integer
    Dim varConfigWert As Variant

    ' 1. Standard-Anspruch aus Sortiment laden
    intStandard = modDaten.GetStandardAnspruch(intArtikelID)

    ' 2. Bereich und Artikelname ermitteln
    strBereich = modDaten.GetMitarbeiterBereich(lngPersonalnummer)
    strArtikelName = modDaten.GetArtikelName(intArtikelID)

    ' 3. Sonderregel prüfen: Innendienst bekommt nur 2 Hemden/Blusen
    If strBereich = "Innendienst" Then
        If strArtikelName = "Hemd" Or strArtikelName = "Bluse" Then
            ' Wert aus Config lesen, falls vorhanden
            varConfigWert = modDaten.GetConfigValue("InnendienstHemdAnspruch")
            If IsNumeric(varConfigWert) And varConfigWert > 0 Then
                intInnendienstAnspruch = CInt(varConfigWert)
            Else
                intInnendienstAnspruch = INNENDIENST_HEMD_DEFAULT
            End If
            BerechneEffektivenAnspruch = intInnendienstAnspruch
            Exit Function
        End If
    End If

    ' 4. Keine Sonderregel: Standard-Anspruch verwenden
    ' Hier könnten später individuelle Abweichungen ergänzt werden
    BerechneEffektivenAnspruch = intStandard
End Function

' ----------------------------------------------------------------------------
' Restanspruch berechnen
' ----------------------------------------------------------------------------
' Berechnet den verbleibenden Anspruch basierend auf Zyklus-Typ:
' - Kalender: Anspruch pro Kalenderjahr
' - Rollierend: Anspruch alle X Jahre ab letzter Ausgabe
' ----------------------------------------------------------------------------
Public Function BerechneRestanspruch(lngPersonalnummer As Long, _
                                     intArtikelID As Integer, _
                                     intJahr As Integer) As Integer
    Dim intEffektiverAnspruch As Integer
    Dim intAusgegeben As Integer
    Dim strZyklusTyp As String
    Dim intZyklusJahre As Integer
    Dim dtLetzteAusgabe As Date

    ' Basis-Werte ermitteln
    intEffektiverAnspruch = BerechneEffektivenAnspruch(lngPersonalnummer, intArtikelID)
    strZyklusTyp = modDaten.GetZyklusTyp(intArtikelID)
    intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)

    If strZyklusTyp = "Kalender" Then
        ' ================================================================
        ' JÄHRLICHER ANSPRUCH (Kalender)
        ' ================================================================
        ' Ausgaben im Kalenderjahr zählen und von Anspruch abziehen
        intAusgegeben = GetAusgabenImJahr(lngPersonalnummer, intArtikelID, intJahr)
        BerechneRestanspruch = intEffektiverAnspruch - intAusgegeben

    Else
        ' ================================================================
        ' ROLLIERENDER ANSPRUCH (z.B. alle 3 Jahre)
        ' ================================================================
        dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)

        If dtLetzteAusgabe = 0 Then
            ' Noch nie ausgegeben -> voller Anspruch (Erstanspruch)
            BerechneRestanspruch = intEffektiverAnspruch

        ElseIf intJahr - Year(dtLetzteAusgabe) >= intZyklusJahre Then
            ' Zyklus abgelaufen -> neuer Anspruch
            BerechneRestanspruch = intEffektiverAnspruch

        Else
            ' Noch im Zyklus -> kein Anspruch
            BerechneRestanspruch = 0
        End If
    End If

    ' Mindestens 0 (keine negativen Ansprüche)
    If BerechneRestanspruch < 0 Then BerechneRestanspruch = 0
End Function

' ----------------------------------------------------------------------------
' Ausgaben im Kalenderjahr zählen
' ----------------------------------------------------------------------------
Public Function GetAusgabenImJahr(lngPersonalnummer As Long, _
                                   intArtikelID As Integer, _
                                   intJahr As Integer) As Integer
    Dim wsAusgaben As Worksheet
    Dim dblSumme As Double

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")

    On Error Resume Next
    dblSumme = Application.WorksheetFunction.SumIfs( _
        wsAusgaben.Range("H:H"), _
        wsAusgaben.Range("C:C"), lngPersonalnummer, _
        wsAusgaben.Range("E:E"), intArtikelID, _
        wsAusgaben.Range("I:I"), intJahr)
    On Error GoTo 0

    GetAusgabenImJahr = CInt(dblSumme)
End Function

' ----------------------------------------------------------------------------
' Letzte Ausgabe eines Artikels für einen Mitarbeiter ermitteln
' ----------------------------------------------------------------------------
Public Function GetLetzteAusgabeDatum(lngPersonalnummer As Long, _
                                      intArtikelID As Integer) As Date
    Dim wsAusgaben As Worksheet
    Dim varMax As Variant

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")

    On Error Resume Next
    varMax = Application.WorksheetFunction.MaxIfs( _
        wsAusgaben.Range("B:B"), _
        wsAusgaben.Range("C:C"), lngPersonalnummer, _
        wsAusgaben.Range("E:E"), intArtikelID)
    On Error GoTo 0

    If IsError(varMax) Or IsEmpty(varMax) Or varMax = 0 Then
        GetLetzteAusgabeDatum = 0
    Else
        GetLetzteAusgabeDatum = CDate(varMax)
    End If
End Function

' ----------------------------------------------------------------------------
' Nächstes Anspruchsjahr für rollierende Artikel ermitteln
' ----------------------------------------------------------------------------
Public Function GetNaechstesAnspruchsjahr(lngPersonalnummer As Long, _
                                          intArtikelID As Integer) As Integer
    Dim dtLetzteAusgabe As Date
    Dim intZyklusJahre As Integer

    dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)
    intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)

    If dtLetzteAusgabe = 0 Then
        ' Noch nie ausgegeben -> sofort berechtigt
        GetNaechstesAnspruchsjahr = Year(Date)
    Else
        GetNaechstesAnspruchsjahr = Year(dtLetzteAusgabe) + intZyklusJahre
    End If
End Function

' ----------------------------------------------------------------------------
' Restanspruch berechnen und auf Restanspruch-Blatt anzeigen
' ----------------------------------------------------------------------------
Public Sub BerechneUndZeigeRestanspruch(lngPersonalnummer As Long, intJahr As Integer)
    Dim wsRest As Worksheet
    Dim wsSortiment As Worksheet
    Dim rngArtikel As Range
    Dim lngRow As Long
    Dim intArtikelID As Integer
    Dim intStandard As Integer
    Dim intEffektiv As Integer
    Dim intAusgegeben As Integer
    Dim intRest As Integer
    Dim strStatus As String
    Dim strZyklusTyp As String
    Dim intZyklusJahre As Integer
    Dim dtLetzteAusgabe As Date
    Dim strMitarbeiterName As String

    Set wsRest = ThisWorkbook.Sheets("Restanspruch")
    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")

    ' Mitarbeitername für Anzeige
    strMitarbeiterName = modDaten.GetMitarbeiterName(lngPersonalnummer)
    If strMitarbeiterName = "" Then
        MsgBox "Mitarbeiter mit Personalnummer " & lngPersonalnummer & " nicht gefunden.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Alte Ergebnisdaten löschen (ab Zeile 8)
    wsRest.Range("A8:F100").ClearContents

    ' Durch alle aktiven Artikel iterieren
    lngRow = 8
    For Each rngArtikel In wsSortiment.ListObjects("tblSortiment").DataBodyRange.Rows
        ' Nur aktive Artikel
        If rngArtikel.Cells(1, 6).Value = "Ja" Then
            intArtikelID = CInt(rngArtikel.Cells(1, 1).Value)

            ' Werte berechnen
            intStandard = modDaten.GetStandardAnspruch(intArtikelID)
            intEffektiv = BerechneEffektivenAnspruch(lngPersonalnummer, intArtikelID)
            intRest = BerechneRestanspruch(lngPersonalnummer, intArtikelID, intJahr)
            strZyklusTyp = modDaten.GetZyklusTyp(intArtikelID)
            intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)

            ' Status ermitteln
            If strZyklusTyp = "Kalender" Then
                ' Jährlicher Artikel
                intAusgegeben = GetAusgabenImJahr(lngPersonalnummer, intArtikelID, intJahr)
                If intRest > 0 Then
                    strStatus = "Verfügbar"
                Else
                    strStatus = "Erschöpft"
                End If
            Else
                ' Rollierender Artikel
                dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)
                intAusgegeben = 0  ' Bei rollierend nicht jahresbezogen

                If dtLetzteAusgabe = 0 Then
                    strStatus = "Verfügbar (noch nie ausgegeben)"
                ElseIf intRest > 0 Then
                    strStatus = "Verfügbar (letzte: " & Year(dtLetzteAusgabe) & ")"
                Else
                    strStatus = "Nächste: " & (Year(dtLetzteAusgabe) + intZyklusJahre)
                End If
            End If

            ' Daten in Blatt schreiben
            With wsRest
                .Cells(lngRow, 1).Value = modDaten.GetArtikelName(intArtikelID)
                .Cells(lngRow, 2).Value = intStandard
                .Cells(lngRow, 3).Value = intEffektiv
                .Cells(lngRow, 4).Value = intAusgegeben
                .Cells(lngRow, 5).Value = intRest
                .Cells(lngRow, 6).Value = strStatus

                ' Farbliche Hervorhebung
                If intRest > 0 Then
                    .Cells(lngRow, 5).Interior.Color = RGB(198, 239, 206)  ' Grün
                Else
                    .Cells(lngRow, 5).Interior.Color = RGB(255, 199, 206)  ' Rot
                End If
            End With

            lngRow = lngRow + 1
        End If
    Next rngArtikel

    ' Info-Zeile mit Mitarbeitername
    wsRest.Range("A6").Value = "Restanspruch für: " & strMitarbeiterName & _
                               " (Personalnr. " & lngPersonalnummer & ") - Jahr " & intJahr
    wsRest.Range("A6").Font.Bold = True
End Sub

' ----------------------------------------------------------------------------
' Übersicht aktualisieren (alle Mitarbeiter, gewähltes Jahr)
' ----------------------------------------------------------------------------
Public Sub AktualisiereUebersicht()
    Dim wsUebersicht As Worksheet
    Dim wsMitarbeiter As Worksheet
    Dim wsSortiment As Worksheet
    Dim rngMitarbeiter As Range
    Dim rngArtikel As Range
    Dim lngRow As Long
    Dim lngCol As Long

    Set wsUebersicht = ThisWorkbook.Sheets("Uebersicht")
    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")

    ' Alte Daten löschen (ab Zeile 6)
    wsUebersicht.Range("A6:Z1000").ClearContents

    ' Header-Zeile für Artikel
    lngCol = 3
    For Each rngArtikel In wsSortiment.ListObjects("tblSortiment").DataBodyRange.Rows
        If rngArtikel.Cells(1, 6).Value = "Ja" Then
            wsUebersicht.Cells(5, lngCol).Value = rngArtikel.Cells(1, 2).Value
            lngCol = lngCol + 1
        End If
    Next rngArtikel

    ' Mitarbeiter eintragen
    lngRow = 6
    For Each rngMitarbeiter In wsMitarbeiter.ListObjects("tblMitarbeiter").DataBodyRange.Rows
        ' Nur aktive Mitarbeiter
        If rngMitarbeiter.Cells(1, 5).Value = "Ja" Then
            ' Personalnummer
            wsUebersicht.Cells(lngRow, 1).Value = rngMitarbeiter.Cells(1, 1).Value

            ' Name (Formel)
            wsUebersicht.Cells(lngRow, 2).Formula = _
                "=IFERROR(VLOOKUP(A" & lngRow & ",tblMitarbeiter,2,FALSE)&"" ""&" & _
                "VLOOKUP(A" & lngRow & ",tblMitarbeiter,3,FALSE),"""")"

            ' Ausgaben pro Artikel (SUMIFS-Formeln)
            lngCol = 3
            For Each rngArtikel In wsSortiment.ListObjects("tblSortiment").DataBodyRange.Rows
                If rngArtikel.Cells(1, 6).Value = "Ja" Then
                    wsUebersicht.Cells(lngRow, lngCol).Formula = _
                        "=SUMIFS(tblAusgaben[Menge]," & _
                        "tblAusgaben[Personalnummer],$A" & lngRow & "," & _
                        "tblAusgaben[ArtikelID]," & rngArtikel.Cells(1, 1).Value & "," & _
                        "tblAusgaben[Kalenderjahr],$B$3)"
                    lngCol = lngCol + 1
                End If
            Next rngArtikel

            lngRow = lngRow + 1
        End If
    Next rngMitarbeiter

    ' Header formatieren
    With wsUebersicht.Range(wsUebersicht.Cells(5, 1), wsUebersicht.Cells(5, lngCol - 1))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Spaltenbreiten anpassen
    wsUebersicht.Columns("A:Z").AutoFit
End Sub
