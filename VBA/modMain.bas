Attribute VB_Name = "modMain"
Option Explicit

' ============================================================================
' modMain - Hauptmodul für Bekleidungsverwaltung
' ============================================================================
' Gemäß App-Architecture/Excel-VBA-Conventions.md
' ============================================================================

Public Const APP_NAME As String = "Bekleidungsverwaltung"
Public Const APP_VERSION As String = "1.0.0"

' ----------------------------------------------------------------------------
' Anwendung initialisieren
' ----------------------------------------------------------------------------
Public Sub InitializeApplication()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Call modHelfer.AktualisiereDropdowns
    Call modHelfer.RefreshNamedRanges

CleanExit:
    Application.ScreenUpdating = True
    MsgBox "Anwendung initialisiert.", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler in InitializeApplication:" & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_NAME
    Resume CleanExit
End Sub

' ----------------------------------------------------------------------------
' Neue Ausgabe hinzufügen (öffnet Eingabedialog)
' ----------------------------------------------------------------------------
Public Sub BtnNeueAusgabe_Click()
    On Error GoTo ErrorHandler

    Dim strDatum As String
    Dim strPersonalnr As String
    Dim strArtikelID As String
    Dim strGroesse As String
    Dim strMenge As String
    Dim strBemerkung As String
    Dim dtDatum As Date
    Dim lngPersonalnr As Long
    Dim intArtikelID As Integer
    Dim intMenge As Integer

    ' Datum abfragen
    strDatum = InputBox("Datum (TT.MM.JJJJ):", APP_NAME, Format(Date, "DD.MM.YYYY"))
    If strDatum = "" Then Exit Sub

    ' Personalnummer abfragen
    strPersonalnr = InputBox("Personalnummer:", APP_NAME)
    If strPersonalnr = "" Then Exit Sub

    ' ArtikelID abfragen
    strArtikelID = InputBox("ArtikelID:" & vbCrLf & _
                            "1 = Hemd" & vbCrLf & _
                            "2 = Bluse" & vbCrLf & _
                            "3 = Polo Shirt" & vbCrLf & _
                            "4 = Hoodie" & vbCrLf & _
                            "5 = Softshelljacke", APP_NAME)
    If strArtikelID = "" Then Exit Sub

    ' Größe abfragen
    strGroesse = InputBox("Größe (XS, S, M, L, XL, XXL):", APP_NAME, "L")
    If strGroesse = "" Then Exit Sub

    ' Menge abfragen
    strMenge = InputBox("Menge:", APP_NAME, "1")
    If strMenge = "" Then Exit Sub

    ' Bemerkung abfragen (optional)
    strBemerkung = InputBox("Bemerkung (optional):", APP_NAME)

    ' Werte konvertieren
    dtDatum = CDate(strDatum)
    lngPersonalnr = CLng(strPersonalnr)
    intArtikelID = CInt(strArtikelID)
    intMenge = CInt(strMenge)

    ' Validierung
    If Not modHelfer.ValidateAusgabe(dtDatum, lngPersonalnr, intArtikelID, intMenge) Then
        Exit Sub
    End If

    ' Ausgabe speichern
    Call modDaten.AddAusgabe(dtDatum, lngPersonalnr, intArtikelID, _
                             strGroesse, intMenge, strBemerkung)

    MsgBox "Ausgabe erfolgreich hinzugefügt!", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler in BtnNeueAusgabe_Click:" & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_NAME
End Sub

' ----------------------------------------------------------------------------
' Übersicht aktualisieren
' ----------------------------------------------------------------------------
Public Sub BtnUebersichtAktualisieren_Click()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Call modBerechnung.AktualisiereUebersicht

CleanExit:
    Application.ScreenUpdating = True
    MsgBox "Übersicht aktualisiert!", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in BtnUebersichtAktualisieren_Click:" & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_NAME
End Sub

' ----------------------------------------------------------------------------
' Restanspruch berechnen
' ----------------------------------------------------------------------------
Public Sub BtnRestanspruchBerechnen_Click()
    On Error GoTo ErrorHandler

    Dim wsRest As Worksheet
    Dim intJahr As Integer
    Dim lngPersonalnr As Long

    Set wsRest = ThisWorkbook.Sheets("Restanspruch")

    ' Werte aus Eingabebereich lesen
    intJahr = CInt(wsRest.Range("B3").Value)
    lngPersonalnr = CLng(wsRest.Range("B4").Value)

    ' Validierung
    If lngPersonalnr = 0 Then
        MsgBox "Bitte geben Sie eine gültige Personalnummer ein.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    If intJahr < 2020 Or intJahr > 2030 Then
        MsgBox "Bitte geben Sie ein gültiges Jahr ein (2020-2030).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Call modBerechnung.BerechneUndZeigeRestanspruch(lngPersonalnr, intJahr)

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in BtnRestanspruchBerechnen_Click:" & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_NAME
End Sub

' ----------------------------------------------------------------------------
' Alle Daten sortieren (Ausgaben nach Datum)
' ----------------------------------------------------------------------------
Public Sub BtnAusgabenSortieren_Click()
    On Error GoTo ErrorHandler

    Dim wsAusgaben As Worksheet
    Dim tblAusgaben As ListObject

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")
    Set tblAusgaben = wsAusgaben.ListObjects("tblAusgaben")

    Application.ScreenUpdating = False

    With tblAusgaben.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tblAusgaben.ListColumns("Datum").Range, _
                        SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

CleanExit:
    Application.ScreenUpdating = True
    MsgBox "Ausgaben nach Datum sortiert!", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Sortieren:" & vbCrLf & Err.Description, _
           vbCritical, APP_NAME
End Sub
