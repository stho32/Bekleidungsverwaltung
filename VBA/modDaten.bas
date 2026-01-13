Attribute VB_Name = "modDaten"
Option Explicit

' ============================================================================
' modDaten - Datenzugriffsschicht
' ============================================================================
' Gemäß App-Architecture/Excel-VBA-Conventions.md
' Alle Datenbankzugriffe (Lesen/Schreiben) erfolgen über dieses Modul
' ============================================================================

' ----------------------------------------------------------------------------
' Mitarbeitername anhand Personalnummer ermitteln
' ----------------------------------------------------------------------------
Public Function GetMitarbeiterName(lngPersonalnummer As Long) As String
    Dim wsMitarbeiter As Worksheet
    Dim rngFound As Range

    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set rngFound = wsMitarbeiter.Range("A:A").Find(lngPersonalnummer, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetMitarbeiterName = wsMitarbeiter.Cells(rngFound.Row, 2).Value & " " & _
                             wsMitarbeiter.Cells(rngFound.Row, 3).Value
    Else
        GetMitarbeiterName = ""
    End If
End Function

' ----------------------------------------------------------------------------
' Bereich des Mitarbeiters ermitteln (Innendienst/Außendienst)
' ----------------------------------------------------------------------------
Public Function GetMitarbeiterBereich(lngPersonalnummer As Long) As String
    Dim wsMitarbeiter As Worksheet
    Dim rngFound As Range

    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set rngFound = wsMitarbeiter.Range("A:A").Find(lngPersonalnummer, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetMitarbeiterBereich = wsMitarbeiter.Cells(rngFound.Row, 6).Value
    Else
        GetMitarbeiterBereich = ""
    End If
End Function

' ----------------------------------------------------------------------------
' Prüfen ob Mitarbeiter aktiv ist
' ----------------------------------------------------------------------------
Public Function IsMitarbeiterAktiv(lngPersonalnummer As Long) As Boolean
    Dim wsMitarbeiter As Worksheet
    Dim rngFound As Range

    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set rngFound = wsMitarbeiter.Range("A:A").Find(lngPersonalnummer, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        IsMitarbeiterAktiv = (wsMitarbeiter.Cells(rngFound.Row, 5).Value = "Ja")
    Else
        IsMitarbeiterAktiv = False
    End If
End Function

' ----------------------------------------------------------------------------
' Artikelname anhand ArtikelID ermitteln
' ----------------------------------------------------------------------------
Public Function GetArtikelName(intArtikelID As Integer) As String
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetArtikelName = wsSortiment.Cells(rngFound.Row, 2).Value
    Else
        GetArtikelName = ""
    End If
End Function

' ----------------------------------------------------------------------------
' Standard-Anspruch für Artikel abrufen
' ----------------------------------------------------------------------------
Public Function GetStandardAnspruch(intArtikelID As Integer) As Integer
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetStandardAnspruch = CInt(wsSortiment.Cells(rngFound.Row, 3).Value)
    Else
        GetStandardAnspruch = 0
    End If
End Function

' ----------------------------------------------------------------------------
' Zyklus-Jahre für Artikel abrufen (1 oder 3)
' ----------------------------------------------------------------------------
Public Function GetZyklusJahre(intArtikelID As Integer) As Integer
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetZyklusJahre = CInt(wsSortiment.Cells(rngFound.Row, 4).Value)
    Else
        GetZyklusJahre = 1
    End If
End Function

' ----------------------------------------------------------------------------
' Zyklus-Typ für Artikel abrufen (Kalender/Rollierend)
' ----------------------------------------------------------------------------
Public Function GetZyklusTyp(intArtikelID As Integer) As String
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetZyklusTyp = wsSortiment.Cells(rngFound.Row, 5).Value
    Else
        GetZyklusTyp = "Kalender"
    End If
End Function

' ----------------------------------------------------------------------------
' Verfügbare Größen für Artikel abrufen
' ----------------------------------------------------------------------------
Public Function GetGroessenFuerArtikel(intArtikelID As Integer) As String
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetGroessenFuerArtikel = wsSortiment.Cells(rngFound.Row, 7).Value
    Else
        GetGroessenFuerArtikel = "S,M,L,XL"
    End If
End Function

' ----------------------------------------------------------------------------
' Neue Ausgabe hinzufügen
' ----------------------------------------------------------------------------
Public Sub AddAusgabe(dtDatum As Date, lngPersonalnummer As Long, _
                      intArtikelID As Integer, strGroesse As String, _
                      intMenge As Integer, strBemerkung As String)
    Dim wsAusgaben As Worksheet
    Dim lngNextRow As Long
    Dim lngNextID As Long

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")

    ' Nächste freie Zeile finden
    lngNextRow = wsAusgaben.Cells(wsAusgaben.Rows.Count, 1).End(xlUp).Row + 1
    lngNextID = GetNextAusgabeID()

    ' Daten eintragen
    With wsAusgaben
        .Cells(lngNextRow, 1).Value = lngNextID                              ' AusgabeID
        .Cells(lngNextRow, 2).Value = dtDatum                                 ' Datum
        .Cells(lngNextRow, 3).Value = lngPersonalnummer                       ' Personalnummer
        .Cells(lngNextRow, 4).Formula = "=IFERROR(VLOOKUP(C" & lngNextRow & _
                                        ",tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(C" & _
                                        lngNextRow & ",tblMitarbeiter,3,FALSE),"""")"  ' Name
        .Cells(lngNextRow, 5).Value = intArtikelID                            ' ArtikelID
        .Cells(lngNextRow, 6).Formula = "=IFERROR(VLOOKUP(E" & lngNextRow & _
                                        ",tblSortiment,2,FALSE),"""")"        ' Artikelname
        .Cells(lngNextRow, 7).Value = strGroesse                              ' Größe
        .Cells(lngNextRow, 8).Value = intMenge                                ' Menge
        .Cells(lngNextRow, 9).Formula = "=YEAR(B" & lngNextRow & ")"          ' Kalenderjahr
        .Cells(lngNextRow, 10).Value = strBemerkung                           ' Bemerkung
    End With
End Sub

' ----------------------------------------------------------------------------
' Nächste AusgabeID ermitteln (Auto-Increment)
' ----------------------------------------------------------------------------
Public Function GetNextAusgabeID() As Long
    Dim wsAusgaben As Worksheet
    Dim lngMaxID As Long

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")

    On Error Resume Next
    lngMaxID = Application.WorksheetFunction.Max(wsAusgaben.Range("A:A"))
    On Error GoTo 0

    If lngMaxID = 0 Then
        GetNextAusgabeID = 1
    Else
        GetNextAusgabeID = lngMaxID + 1
    End If
End Function

' ----------------------------------------------------------------------------
' Config-Wert abrufen
' ----------------------------------------------------------------------------
Public Function GetConfigValue(strParameter As String) As Variant
    Dim wsConfig As Worksheet
    Dim varResult As Variant

    Set wsConfig = ThisWorkbook.Sheets("Config")

    On Error Resume Next
    varResult = Application.WorksheetFunction.VLookup( _
                    strParameter, wsConfig.Range("A:B"), 2, False)
    On Error GoTo 0

    If IsError(varResult) Then
        GetConfigValue = ""
    Else
        GetConfigValue = varResult
    End If
End Function

' ----------------------------------------------------------------------------
' Liste aller aktiven Mitarbeiter abrufen
' ----------------------------------------------------------------------------
Public Function GetAktiveMitarbeiter() As Collection
    Dim wsMitarbeiter As Worksheet
    Dim tblMitarbeiter As ListObject
    Dim rngRow As Range
    Dim colResult As New Collection

    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set tblMitarbeiter = wsMitarbeiter.ListObjects("tblMitarbeiter")

    For Each rngRow In tblMitarbeiter.DataBodyRange.Rows
        If rngRow.Cells(1, 5).Value = "Ja" Then
            colResult.Add rngRow.Cells(1, 1).Value & " - " & _
                          rngRow.Cells(1, 2).Value & " " & _
                          rngRow.Cells(1, 3).Value
        End If
    Next rngRow

    Set GetAktiveMitarbeiter = colResult
End Function

' ----------------------------------------------------------------------------
' Liste aller aktiven Artikel abrufen
' ----------------------------------------------------------------------------
Public Function GetAktiveArtikel() As Collection
    Dim wsSortiment As Worksheet
    Dim tblSortiment As ListObject
    Dim rngRow As Range
    Dim colResult As New Collection

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set tblSortiment = wsSortiment.ListObjects("tblSortiment")

    For Each rngRow In tblSortiment.DataBodyRange.Rows
        If rngRow.Cells(1, 6).Value = "Ja" Then
            colResult.Add rngRow.Cells(1, 1).Value & " - " & _
                          rngRow.Cells(1, 2).Value
        End If
    Next rngRow

    Set GetAktiveArtikel = colResult
End Function
