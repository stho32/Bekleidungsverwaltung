Attribute VB_Name = "modHelfer"
Option Explicit

' ============================================================================
' modHelfer - Hilfsfunktionen
' ============================================================================
' Gemäß App-Architecture/Excel-VBA-Conventions.md
' Enthält Validierung, Formatierung und andere Hilfsfunktionen
' ============================================================================

' ----------------------------------------------------------------------------
' Ausgabe validieren
' ----------------------------------------------------------------------------
' Prüft alle Eingabewerte vor dem Speichern einer Ausgabe
' Gibt True zurück wenn alle Werte gültig sind
' ----------------------------------------------------------------------------
Public Function ValidateAusgabe(dtDatum As Date, _
                                lngPersonalnummer As Long, _
                                intArtikelID As Integer, _
                                intMenge As Integer) As Boolean
    ValidateAusgabe = True

    ' Datum nicht in Zukunft
    If dtDatum > Date Then
        MsgBox "Datum darf nicht in der Zukunft liegen.", _
               vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    ' Datum nicht zu weit in der Vergangenheit (max 2 Jahre)
    If dtDatum < DateAdd("yyyy", -2, Date) Then
        Dim intAntwort As Integer
        intAntwort = MsgBox("Das Datum liegt mehr als 2 Jahre in der Vergangenheit." & vbCrLf & _
                            "Möchten Sie trotzdem fortfahren?", _
                            vbQuestion + vbYesNo, APP_NAME)
        If intAntwort = vbNo Then
            ValidateAusgabe = False
            Exit Function
        End If
    End If

    ' Mitarbeiter existiert
    If modDaten.GetMitarbeiterName(lngPersonalnummer) = "" Then
        MsgBox "Mitarbeiter mit Personalnummer " & lngPersonalnummer & " nicht gefunden.", _
               vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    ' Mitarbeiter ist aktiv
    If Not modDaten.IsMitarbeiterAktiv(lngPersonalnummer) Then
        MsgBox "Mitarbeiter mit Personalnummer " & lngPersonalnummer & " ist nicht aktiv.", _
               vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    ' Artikel existiert
    If modDaten.GetArtikelName(intArtikelID) = "" Then
        MsgBox "Artikel mit ID " & intArtikelID & " nicht gefunden.", _
               vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    ' Menge positiv
    If intMenge <= 0 Then
        MsgBox "Menge muss größer als 0 sein.", _
               vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    ' Menge nicht zu hoch (max 10)
    If intMenge > 10 Then
        MsgBox "Menge darf maximal 10 sein.", _
               vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    ' Optional: Restanspruch prüfen
    Dim intRest As Integer
    intRest = modBerechnung.BerechneRestanspruch(lngPersonalnummer, intArtikelID, Year(dtDatum))
    If intMenge > intRest Then
        Dim intAntwort2 As Integer
        intAntwort2 = MsgBox("Die angeforderte Menge (" & intMenge & ") übersteigt den " & _
                             "Restanspruch (" & intRest & ")." & vbCrLf & vbCrLf & _
                             "Möchten Sie trotzdem fortfahren?", _
                             vbQuestion + vbYesNo, APP_NAME)
        If intAntwort2 = vbNo Then
            ValidateAusgabe = False
            Exit Function
        End If
    End If
End Function

' ----------------------------------------------------------------------------
' Benannte Bereiche aktualisieren
' ----------------------------------------------------------------------------
' Aktualisiert dynamische benannte Bereiche nach Datenänderungen
' ----------------------------------------------------------------------------
Public Sub RefreshNamedRanges()
    On Error Resume Next

    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' Die benannten Bereiche werden automatisch durch Excel-Tabellen
    ' (ListObjects) verwaltet, daher ist keine manuelle Aktualisierung nötig

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Dropdowns aktualisieren
' ----------------------------------------------------------------------------
' Aktualisiert Datenvalidierungen für Dropdown-Felder
' ----------------------------------------------------------------------------
Public Sub AktualisiereDropdowns()
    On Error Resume Next

    ' Mitarbeiter-Dropdown im Restanspruch-Blatt
    Dim wsRest As Worksheet
    Set wsRest = ThisWorkbook.Sheets("Restanspruch")

    ' Die Dropdowns werden durch Excel-Tabellen automatisch aktualisiert
    ' Hier könnten bei Bedarf weitere Anpassungen erfolgen

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Größen-Array aus String erstellen
' ----------------------------------------------------------------------------
Public Function ParseGroessen(strGroessen As String) As Variant
    If strGroessen = "" Then
        ParseGroessen = Array("S", "M", "L", "XL")
    Else
        ParseGroessen = Split(strGroessen, ",")
    End If
End Function

' ----------------------------------------------------------------------------
' Performance-Wrapper: Optimierungen aktivieren
' ----------------------------------------------------------------------------
Public Sub StartPerformanceMode()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

' ----------------------------------------------------------------------------
' Performance-Wrapper: Optimierungen deaktivieren
' ----------------------------------------------------------------------------
Public Sub EndPerformanceMode()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ----------------------------------------------------------------------------
' Aktivität protokollieren (optional)
' ----------------------------------------------------------------------------
Public Sub LogActivity(strMessage As String)
    ' Kann bei Bedarf implementiert werden
    ' z.B. Schreiben in ein Log-Blatt oder externe Datei
    Debug.Print Format(Now, "YYYY-MM-DD HH:MM:SS") & " - " & strMessage
End Sub

' ----------------------------------------------------------------------------
' Prüfen ob Blatt existiert
' ----------------------------------------------------------------------------
Public Function SheetExists(strSheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(strSheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing
End Function

' ----------------------------------------------------------------------------
' Prüfen ob Tabelle (ListObject) existiert
' ----------------------------------------------------------------------------
Public Function TableExists(strTableName As String) As Boolean
    Dim lo As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set lo = ws.ListObjects(strTableName)
        If Not lo Is Nothing Then
            TableExists = True
            Exit Function
        End If
    Next ws
    On Error GoTo 0

    TableExists = False
End Function

' ----------------------------------------------------------------------------
' Letzte Zeile einer Spalte finden
' ----------------------------------------------------------------------------
Public Function GetLastRow(ws As Worksheet, intColumn As Integer) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, intColumn).End(xlUp).Row
End Function

' ----------------------------------------------------------------------------
' Letzte Spalte einer Zeile finden
' ----------------------------------------------------------------------------
Public Function GetLastColumn(ws As Worksheet, intRow As Integer) As Long
    GetLastColumn = ws.Cells(intRow, ws.Columns.Count).End(xlToLeft).Column
End Function

' ----------------------------------------------------------------------------
' Zelle formatieren basierend auf Wert
' ----------------------------------------------------------------------------
Public Sub FormatiereCelleNachWert(rngCell As Range, _
                                    Optional dblGrenze As Double = 0)
    If IsNumeric(rngCell.Value) Then
        If rngCell.Value > dblGrenze Then
            rngCell.Interior.Color = RGB(198, 239, 206)  ' Grün
        ElseIf rngCell.Value = dblGrenze Then
            rngCell.Interior.Color = RGB(255, 235, 156)  ' Gelb
        Else
            rngCell.Interior.Color = RGB(255, 199, 206)  ' Rot
        End If
    End If
End Sub

' ----------------------------------------------------------------------------
' Alle Daten in einer Tabelle zählen
' ----------------------------------------------------------------------------
Public Function CountTableRows(strTableName As String) As Long
    Dim lo As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set lo = ws.ListObjects(strTableName)
        If Not lo Is Nothing Then
            If lo.DataBodyRange Is Nothing Then
                CountTableRows = 0
            Else
                CountTableRows = lo.DataBodyRange.Rows.Count
            End If
            Exit Function
        End If
    Next ws
    On Error GoTo 0

    CountTableRows = 0
End Function
