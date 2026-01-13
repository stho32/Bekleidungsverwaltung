' ============================================================================
' Bekleidungsverwaltung Setup Script
' ============================================================================
' Dieses VBScript erstellt die komplette Excel-Lösung für die
' Bekleidungskontingent-Verwaltung.
'
' Ausführung: Doppelklick auf diese Datei oder via Kommandozeile:
'             cscript Bekleidungsverwaltung_Setup.vbs
' ============================================================================

Option Explicit

Dim objExcel, objWorkbook, objSheet
Dim strPath, strFileName

' Pfad zur Zieldatei
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strFileName = strPath & "\Bekleidungsverwaltung.xlsm"

' Excel starten
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.DisplayAlerts = False

' Neue Arbeitsmappe erstellen
Set objWorkbook = objExcel.Workbooks.Add

' Standard-Blätter entfernen und neue anlegen
Call ErstelleTabellenBlaetter(objWorkbook)

' Spaltenstrukturen anlegen
Call ErstelleMitarbeiterBlatt(objWorkbook.Sheets("Mitarbeiter"))
Call ErstelleSortimentBlatt(objWorkbook.Sheets("Sortiment"))
Call ErstelleAusgabenBlatt(objWorkbook.Sheets("Ausgaben"))
Call ErstelleUebersichtBlatt(objWorkbook.Sheets("Uebersicht"))
Call ErstelleRestanspruchBlatt(objWorkbook.Sheets("Restanspruch"))
Call ErstelleConfigBlatt(objWorkbook.Sheets("Config"))

' Benannte Bereiche erstellen
Call ErstelleBenannteBereich(objWorkbook)

' VBA-Code einfügen
Call FuegeVBACodeEin(objWorkbook)

' Speichern als .xlsm (mit Makros)
objWorkbook.SaveAs strFileName, 52 ' 52 = xlOpenXMLWorkbookMacroEnabled

objExcel.DisplayAlerts = True

MsgBox "Bekleidungsverwaltung.xlsm wurde erfolgreich erstellt!" & vbCrLf & vbCrLf & _
       "Pfad: " & strFileName, vbInformation, "Setup abgeschlossen"

' ============================================================================
' TABELLENBLÄTTER ERSTELLEN
' ============================================================================
Sub ErstelleTabellenBlaetter(wb)
    Dim arrSheets, i
    arrSheets = Array("Mitarbeiter", "Sortiment", "Ausgaben", "Uebersicht", "Restanspruch", "Config")

    ' Vorhandene Blätter entfernen (außer erstem)
    Do While wb.Sheets.Count > 1
        wb.Sheets(wb.Sheets.Count).Delete
    Loop

    ' Erstes Blatt umbenennen
    wb.Sheets(1).Name = arrSheets(0)

    ' Weitere Blätter hinzufügen
    For i = 1 To UBound(arrSheets)
        wb.Sheets.Add , wb.Sheets(wb.Sheets.Count)
        wb.Sheets(wb.Sheets.Count).Name = arrSheets(i)
    Next
End Sub

' ============================================================================
' MITARBEITER-BLATT
' ============================================================================
Sub ErstelleMitarbeiterBlatt(ws)
    ' Überschriften
    ws.Range("A1").Value = "Personalnummer"
    ws.Range("B1").Value = "Nachname"
    ws.Range("C1").Value = "Vorname"
    ws.Range("D1").Value = "Eintrittsdatum"
    ws.Range("E1").Value = "Aktiv"
    ws.Range("F1").Value = "Bereich"
    ws.Range("G1").Value = "Abteilung"

    ' Formatierung Überschriften
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Beispieldaten
    ws.Range("A2").Value = 1001
    ws.Range("B2").Value = "Müller"
    ws.Range("C2").Value = "Hans"
    ws.Range("D2").Value = DateSerial(2020, 3, 15)
    ws.Range("E2").Value = "Ja"
    ws.Range("F2").Value = "Außendienst"
    ws.Range("G2").Value = "Vertrieb"

    ws.Range("A3").Value = 1002
    ws.Range("B3").Value = "Schmidt"
    ws.Range("C3").Value = "Anna"
    ws.Range("D3").Value = DateSerial(2019, 7, 1)
    ws.Range("E3").Value = "Ja"
    ws.Range("F3").Value = "Innendienst"
    ws.Range("G3").Value = "Buchhaltung"

    ws.Range("A4").Value = 1003
    ws.Range("B4").Value = "Weber"
    ws.Range("C4").Value = "Thomas"
    ws.Range("D4").Value = DateSerial(2021, 1, 10)
    ws.Range("E4").Value = "Ja"
    ws.Range("F4").Value = "Außendienst"
    ws.Range("G4").Value = "Technik"

    ' Datenvalidierung für Aktiv (Ja/Nein)
    With ws.Range("E2:E1000").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="Ja,Nein"
        .ShowDropDown = False
    End With

    ' Datenvalidierung für Bereich
    With ws.Range("F2:F1000").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="Außendienst,Innendienst"
        .ShowDropDown = False
    End With

    ' Spaltenbreiten
    ws.Columns("A:G").AutoFit

    ' Als Tabelle formatieren
    ws.ListObjects.Add(1, ws.Range("A1:G4"), , 1).Name = "tblMitarbeiter"
End Sub

' ============================================================================
' SORTIMENT-BLATT
' ============================================================================
Sub ErstelleSortimentBlatt(ws)
    ' Überschriften
    ws.Range("A1").Value = "ArtikelID"
    ws.Range("B1").Value = "Artikelname"
    ws.Range("C1").Value = "AnspruchMenge"
    ws.Range("D1").Value = "ZyklusJahre"
    ws.Range("E1").Value = "ZyklusTyp"
    ws.Range("F1").Value = "Aktiv"
    ws.Range("G1").Value = "Groessen"

    ' Formatierung Überschriften
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Beispieldaten - Artikel
    ' Hemd
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "Hemd"
    ws.Range("C2").Value = 4
    ws.Range("D2").Value = 1
    ws.Range("E2").Value = "Kalender"
    ws.Range("F2").Value = "Ja"
    ws.Range("G2").Value = "S,M,L,XL,XXL"

    ' Bluse
    ws.Range("A3").Value = 2
    ws.Range("B3").Value = "Bluse"
    ws.Range("C3").Value = 4
    ws.Range("D3").Value = 1
    ws.Range("E3").Value = "Kalender"
    ws.Range("F3").Value = "Ja"
    ws.Range("G3").Value = "XS,S,M,L,XL"

    ' Polo Shirt
    ws.Range("A4").Value = 3
    ws.Range("B4").Value = "Polo Shirt"
    ws.Range("C4").Value = 2
    ws.Range("D4").Value = 1
    ws.Range("E4").Value = "Kalender"
    ws.Range("F4").Value = "Ja"
    ws.Range("G4").Value = "S,M,L,XL,XXL"

    ' Hoodie
    ws.Range("A5").Value = 4
    ws.Range("B5").Value = "Hoodie"
    ws.Range("C5").Value = 1
    ws.Range("D5").Value = 3
    ws.Range("E5").Value = "Rollierend"
    ws.Range("F5").Value = "Ja"
    ws.Range("G5").Value = "S,M,L,XL,XXL"

    ' Softshelljacke
    ws.Range("A6").Value = 5
    ws.Range("B6").Value = "Softshelljacke"
    ws.Range("C6").Value = 1
    ws.Range("D6").Value = 3
    ws.Range("E6").Value = "Rollierend"
    ws.Range("F6").Value = "Ja"
    ws.Range("G6").Value = "S,M,L,XL,XXL"

    ' Datenvalidierung für ZyklusTyp
    With ws.Range("E2:E100").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="Kalender,Rollierend"
        .ShowDropDown = False
    End With

    ' Datenvalidierung für Aktiv
    With ws.Range("F2:F100").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="Ja,Nein"
        .ShowDropDown = False
    End With

    ' Spaltenbreiten
    ws.Columns("A:G").AutoFit

    ' Als Tabelle formatieren
    ws.ListObjects.Add(1, ws.Range("A1:G6"), , 1).Name = "tblSortiment"
End Sub

' ============================================================================
' AUSGABEN-BLATT
' ============================================================================
Sub ErstelleAusgabenBlatt(ws)
    ' Überschriften
    ws.Range("A1").Value = "AusgabeID"
    ws.Range("B1").Value = "Datum"
    ws.Range("C1").Value = "Personalnummer"
    ws.Range("D1").Value = "MitarbeiterName"
    ws.Range("E1").Value = "ArtikelID"
    ws.Range("F1").Value = "Artikelname"
    ws.Range("G1").Value = "Groesse"
    ws.Range("H1").Value = "Menge"
    ws.Range("I1").Value = "Kalenderjahr"
    ws.Range("J1").Value = "Bemerkung"

    ' Formatierung Überschriften
    With ws.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Beispieldaten
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = DateSerial(2025, 1, 15)
    ws.Range("C2").Value = 1001
    ws.Range("D2").Formula = "=IFERROR(VLOOKUP(C2,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(C2,tblMitarbeiter,3,FALSE),"""")"
    ws.Range("E2").Value = 1
    ws.Range("F2").Formula = "=IFERROR(VLOOKUP(E2,tblSortiment,2,FALSE),"""")"
    ws.Range("G2").Value = "L"
    ws.Range("H2").Value = 2
    ws.Range("I2").Formula = "=YEAR(B2)"
    ws.Range("J2").Value = ""

    ws.Range("A3").Value = 2
    ws.Range("B3").Value = DateSerial(2025, 1, 15)
    ws.Range("C3").Value = 1002
    ws.Range("D3").Formula = "=IFERROR(VLOOKUP(C3,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(C3,tblMitarbeiter,3,FALSE),"""")"
    ws.Range("E3").Value = 2
    ws.Range("F3").Formula = "=IFERROR(VLOOKUP(E3,tblSortiment,2,FALSE),"""")"
    ws.Range("G3").Value = "M"
    ws.Range("H3").Value = 1
    ws.Range("I3").Formula = "=YEAR(B3)"
    ws.Range("J3").Value = ""

    ' Datumsformat
    ws.Columns("B").NumberFormat = "DD.MM.YYYY"

    ' Spaltenbreiten
    ws.Columns("A:J").AutoFit

    ' Als Tabelle formatieren
    ws.ListObjects.Add(1, ws.Range("A1:J3"), , 1).Name = "tblAusgaben"
End Sub

' ============================================================================
' ÜBERSICHT-BLATT
' ============================================================================
Sub ErstelleUebersichtBlatt(ws)
    ' Überschrift
    ws.Range("A1").Value = "Ausgabenübersicht nach Jahr"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Jahr-Auswahl
    ws.Range("A3").Value = "Jahr:"
    ws.Range("B3").Value = Year(Now)

    ' Datenvalidierung für Jahr
    With ws.Range("B3").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="2024,2025,2026,2027,2028,2029,2030"
        .ShowDropDown = False
    End With

    ' Tabellen-Header
    ws.Range("A5").Value = "Personalnummer"
    ws.Range("B5").Value = "Name"
    ws.Range("C5").Value = "Hemd"
    ws.Range("D5").Value = "Bluse"
    ws.Range("E5").Value = "Polo Shirt"
    ws.Range("F5").Value = "Hoodie"
    ws.Range("G5").Value = "Softshelljacke"

    With ws.Range("A5:G5")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Beispielformeln (werden später durch VBA aktualisiert)
    ws.Range("A6").Value = 1001
    ws.Range("B6").Formula = "=IFERROR(VLOOKUP(A6,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(A6,tblMitarbeiter,3,FALSE),"""")"
    ws.Range("C6").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A6,tblAusgaben[ArtikelID],1,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("D6").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A6,tblAusgaben[ArtikelID],2,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("E6").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A6,tblAusgaben[ArtikelID],3,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("F6").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A6,tblAusgaben[ArtikelID],4,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("G6").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A6,tblAusgaben[ArtikelID],5,tblAusgaben[Kalenderjahr],$B$3)"

    ws.Range("A7").Value = 1002
    ws.Range("B7").Formula = "=IFERROR(VLOOKUP(A7,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(A7,tblMitarbeiter,3,FALSE),"""")"
    ws.Range("C7").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A7,tblAusgaben[ArtikelID],1,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("D7").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A7,tblAusgaben[ArtikelID],2,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("E7").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A7,tblAusgaben[ArtikelID],3,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("F7").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A7,tblAusgaben[ArtikelID],4,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("G7").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A7,tblAusgaben[ArtikelID],5,tblAusgaben[Kalenderjahr],$B$3)"

    ws.Range("A8").Value = 1003
    ws.Range("B8").Formula = "=IFERROR(VLOOKUP(A8,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(A8,tblMitarbeiter,3,FALSE),"""")"
    ws.Range("C8").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A8,tblAusgaben[ArtikelID],1,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("D8").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A8,tblAusgaben[ArtikelID],2,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("E8").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A8,tblAusgaben[ArtikelID],3,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("F8").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A8,tblAusgaben[ArtikelID],4,tblAusgaben[Kalenderjahr],$B$3)"
    ws.Range("G8").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A8,tblAusgaben[ArtikelID],5,tblAusgaben[Kalenderjahr],$B$3)"

    ' Spaltenbreiten
    ws.Columns("A:G").AutoFit

    ' Button-Platzhalter Hinweis
    ws.Range("A10").Value = "Hinweis: Klicken Sie auf den Button 'Übersicht aktualisieren' um alle Mitarbeiter anzuzeigen."
    ws.Range("A10").Font.Italic = True
End Sub

' ============================================================================
' RESTANSPRUCH-BLATT
' ============================================================================
Sub ErstelleRestanspruchBlatt(ws)
    ' Überschrift
    ws.Range("A1").Value = "Restanspruch-Abfrage"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Eingabebereich
    ws.Range("A3").Value = "Jahr:"
    ws.Range("B3").Value = Year(Now)

    ws.Range("A4").Value = "Mitarbeiter (Personalnr.):"
    ws.Range("B4").Value = 1001

    ' Formatierung Eingabebereich
    With ws.Range("B3:B4")
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = 1
    End With

    ' Datenvalidierung für Jahr
    With ws.Range("B3").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="2024,2025,2026,2027,2028,2029,2030"
        .ShowDropDown = False
    End With

    ' Ergebnis-Header
    ws.Range("A7").Value = "Artikel"
    ws.Range("B7").Value = "Standard"
    ws.Range("C7").Value = "Effektiv"
    ws.Range("D7").Value = "Ausgegeben"
    ws.Range("E7").Value = "Rest"
    ws.Range("F7").Value = "Status"

    With ws.Range("A7:F7")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Beispiel-Datenzeilen (werden durch VBA befüllt)
    ws.Range("A8").Value = "Hemd"
    ws.Range("B8").Value = 4
    ws.Range("C8").Value = 4
    ws.Range("D8").Value = 2
    ws.Range("E8").Value = 2
    ws.Range("F8").Value = "Verfügbar"

    ws.Range("A9").Value = "Bluse"
    ws.Range("B9").Value = 4
    ws.Range("C9").Value = 4
    ws.Range("D9").Value = 0
    ws.Range("E9").Value = 4
    ws.Range("F9").Value = "Verfügbar"

    ws.Range("A10").Value = "Polo Shirt"
    ws.Range("B10").Value = 2
    ws.Range("C10").Value = 2
    ws.Range("D10").Value = 0
    ws.Range("E10").Value = 2
    ws.Range("F10").Value = "Verfügbar"

    ws.Range("A11").Value = "Hoodie"
    ws.Range("B11").Value = 1
    ws.Range("C11").Value = 1
    ws.Range("D11").Value = 0
    ws.Range("E11").Value = 1
    ws.Range("F11").Value = "Verfügbar (3-Jahres-Zyklus)"

    ws.Range("A12").Value = "Softshelljacke"
    ws.Range("B12").Value = 1
    ws.Range("C12").Value = 1
    ws.Range("D12").Value = 0
    ws.Range("E12").Value = 1
    ws.Range("F12").Value = "Verfügbar (3-Jahres-Zyklus)"

    ' Spaltenbreiten
    ws.Columns("A:F").AutoFit

    ' Hinweis
    ws.Range("A14").Value = "Hinweis: Klicken Sie auf 'Berechnen' um den aktuellen Restanspruch anzuzeigen."
    ws.Range("A14").Font.Italic = True

    ws.Range("A15").Value = "Bei 3-Jahres-Artikeln wird geprüft, ob seit der letzten Ausgabe 3 Jahre vergangen sind."
    ws.Range("A15").Font.Italic = True
End Sub

' ============================================================================
' CONFIG-BLATT
' ============================================================================
Sub ErstelleConfigBlatt(ws)
    ' Überschrift
    ws.Range("A1").Value = "Systemkonfiguration"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Konfigurationsparameter
    ws.Range("A3").Value = "Parameter"
    ws.Range("B3").Value = "Wert"
    ws.Range("C3").Value = "Beschreibung"

    With ws.Range("A3:C3")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Range("A4").Value = "StartJahr"
    ws.Range("B4").Value = 2025
    ws.Range("C4").Value = "Erstes Jahr für Datenerfassung"

    ws.Range("A5").Value = "MaxZeilenAusgaben"
    ws.Range("B5").Value = 10000
    ws.Range("C5").Value = "Maximale Anzahl Ausgabe-Einträge"

    ws.Range("A6").Value = "InnendienstHemdAnspruch"
    ws.Range("B6").Value = 2
    ws.Range("C6").Value = "Reduzierter Hemd/Blusen-Anspruch für Innendienst"

    ws.Range("A7").Value = "AppVersion"
    ws.Range("B7").Value = "1.0.0"
    ws.Range("C7").Value = "Version der Anwendung"

    ' Spaltenbreiten
    ws.Columns("A:C").AutoFit

    ' Als Tabelle formatieren
    ws.ListObjects.Add(1, ws.Range("A3:C7"), , 1).Name = "tblConfig"
End Sub

' ============================================================================
' BENANNTE BEREICHE
' ============================================================================
Sub ErstelleBenannteBereich(wb)
    ' Dynamische benannte Bereiche werden später über VBA erstellt
    ' da sie von den Tabellen abhängen
End Sub

' ============================================================================
' VBA-CODE EINFÜGEN
' ============================================================================
Sub FuegeVBACodeEin(wb)
    Dim objVBProject, objModule

    On Error Resume Next
    Set objVBProject = wb.VBProject

    If Err.Number <> 0 Then
        MsgBox "VBA-Code konnte nicht eingefügt werden." & vbCrLf & _
               "Bitte aktivieren Sie 'Zugriff auf das VBA-Projektobjektmodell vertrauen'" & vbCrLf & _
               "unter Datei > Optionen > Trust Center > Einstellungen für das Trust Center > Makroeinstellungen", _
               vbExclamation, "VBA-Zugriff"
        Exit Sub
    End If
    On Error GoTo 0

    ' modMain
    Set objModule = objVBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
    objModule.Name = "modMain"
    objModule.CodeModule.AddFromString GetModMainCode()

    ' modDaten
    Set objModule = objVBProject.VBComponents.Add(1)
    objModule.Name = "modDaten"
    objModule.CodeModule.AddFromString GetModDatenCode()

    ' modBerechnung
    Set objModule = objVBProject.VBComponents.Add(1)
    objModule.Name = "modBerechnung"
    objModule.CodeModule.AddFromString GetModBerechnungCode()

    ' modHelfer
    Set objModule = objVBProject.VBComponents.Add(1)
    objModule.Name = "modHelfer"
    objModule.CodeModule.AddFromString GetModHelferCode()
End Sub

' ============================================================================
' VBA-CODE: modMain
' ============================================================================
Function GetModMainCode()
    Dim strCode
    strCode = "Option Explicit" & vbCrLf & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf
    strCode = strCode & "' modMain - Hauptmodul für Bekleidungsverwaltung" & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf & vbCrLf
    strCode = strCode & "Public Const APP_NAME As String = ""Bekleidungsverwaltung""" & vbCrLf
    strCode = strCode & "Public Const APP_VERSION As String = ""1.0.0""" & vbCrLf & vbCrLf
    strCode = strCode & "' Anwendung initialisieren" & vbCrLf
    strCode = strCode & "Public Sub InitializeApplication()" & vbCrLf
    strCode = strCode & "    On Error GoTo ErrorHandler" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Call modHelfer.AktualisiereDropdowns" & vbCrLf
    strCode = strCode & "    Call modHelfer.RefreshNamedRanges" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    MsgBox ""Anwendung initialisiert."", vbInformation, APP_NAME" & vbCrLf
    strCode = strCode & "    Exit Sub" & vbCrLf & vbCrLf
    strCode = strCode & "ErrorHandler:" & vbCrLf
    strCode = strCode & "    MsgBox ""Fehler in InitializeApplication: "" & Err.Description, vbCritical, APP_NAME" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf & vbCrLf
    strCode = strCode & "' Neue Ausgabe hinzufügen (öffnet Eingabedialog)" & vbCrLf
    strCode = strCode & "Public Sub BtnNeueAusgabe_Click()" & vbCrLf
    strCode = strCode & "    On Error GoTo ErrorHandler" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Dim strDatum As String, strPersonalnr As String" & vbCrLf
    strCode = strCode & "    Dim strArtikelID As String, strGroesse As String" & vbCrLf
    strCode = strCode & "    Dim strMenge As String, strBemerkung As String" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    strDatum = InputBox(""Datum (TT.MM.JJJJ):"", APP_NAME, Format(Date, ""DD.MM.YYYY""))" & vbCrLf
    strCode = strCode & "    If strDatum = """" Then Exit Sub" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    strPersonalnr = InputBox(""Personalnummer:"", APP_NAME)" & vbCrLf
    strCode = strCode & "    If strPersonalnr = """" Then Exit Sub" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    strArtikelID = InputBox(""ArtikelID (1=Hemd, 2=Bluse, 3=Polo, 4=Hoodie, 5=Softshell):"", APP_NAME)" & vbCrLf
    strCode = strCode & "    If strArtikelID = """" Then Exit Sub" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    strGroesse = InputBox(""Größe (S, M, L, XL, XXL):"", APP_NAME, ""L"")" & vbCrLf
    strCode = strCode & "    If strGroesse = """" Then Exit Sub" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    strMenge = InputBox(""Menge:"", APP_NAME, ""1"")" & vbCrLf
    strCode = strCode & "    If strMenge = """" Then Exit Sub" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    strBemerkung = InputBox(""Bemerkung (optional):"", APP_NAME)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Call modDaten.AddAusgabe(CDate(strDatum), CLng(strPersonalnr), _" & vbCrLf
    strCode = strCode & "                             CInt(strArtikelID), strGroesse, _" & vbCrLf
    strCode = strCode & "                             CInt(strMenge), strBemerkung)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    MsgBox ""Ausgabe erfolgreich hinzugefügt!"", vbInformation, APP_NAME" & vbCrLf
    strCode = strCode & "    Exit Sub" & vbCrLf & vbCrLf
    strCode = strCode & "ErrorHandler:" & vbCrLf
    strCode = strCode & "    MsgBox ""Fehler: "" & Err.Description, vbCritical, APP_NAME" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf & vbCrLf
    strCode = strCode & "' Übersicht aktualisieren" & vbCrLf
    strCode = strCode & "Public Sub BtnUebersichtAktualisieren_Click()" & vbCrLf
    strCode = strCode & "    On Error GoTo ErrorHandler" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Call modBerechnung.AktualisiereUebersicht" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    MsgBox ""Übersicht aktualisiert!"", vbInformation, APP_NAME" & vbCrLf
    strCode = strCode & "    Exit Sub" & vbCrLf & vbCrLf
    strCode = strCode & "ErrorHandler:" & vbCrLf
    strCode = strCode & "    MsgBox ""Fehler: "" & Err.Description, vbCritical, APP_NAME" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf & vbCrLf
    strCode = strCode & "' Restanspruch berechnen" & vbCrLf
    strCode = strCode & "Public Sub BtnRestanspruchBerechnen_Click()" & vbCrLf
    strCode = strCode & "    On Error GoTo ErrorHandler" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Dim wsRest As Worksheet" & vbCrLf
    strCode = strCode & "    Dim intJahr As Integer, lngPersonalnr As Long" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsRest = ThisWorkbook.Sheets(""Restanspruch"")" & vbCrLf
    strCode = strCode & "    intJahr = CInt(wsRest.Range(""B3"").Value)" & vbCrLf
    strCode = strCode & "    lngPersonalnr = CLng(wsRest.Range(""B4"").Value)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Call modBerechnung.BerechneUndZeigeRestanspruch(lngPersonalnr, intJahr)" & vbCrLf
    strCode = strCode & "    Exit Sub" & vbCrLf & vbCrLf
    strCode = strCode & "ErrorHandler:" & vbCrLf
    strCode = strCode & "    MsgBox ""Fehler: "" & Err.Description, vbCritical, APP_NAME" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf

    GetModMainCode = strCode
End Function

' ============================================================================
' VBA-CODE: modDaten
' ============================================================================
Function GetModDatenCode()
    Dim strCode
    strCode = "Option Explicit" & vbCrLf & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf
    strCode = strCode & "' modDaten - Datenzugriffsschicht" & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf & vbCrLf
    strCode = strCode & "' Mitarbeitername anhand Personalnummer ermitteln" & vbCrLf
    strCode = strCode & "Public Function GetMitarbeiterName(lngPersonalnummer As Long) As String" & vbCrLf
    strCode = strCode & "    Dim wsMitarbeiter As Worksheet" & vbCrLf
    strCode = strCode & "    Dim rngFound As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsMitarbeiter = ThisWorkbook.Sheets(""Mitarbeiter"")" & vbCrLf
    strCode = strCode & "    Set rngFound = wsMitarbeiter.Range(""A:A"").Find(lngPersonalnummer, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If Not rngFound Is Nothing Then" & vbCrLf
    strCode = strCode & "        GetMitarbeiterName = wsMitarbeiter.Cells(rngFound.Row, 2).Value & "" "" & _" & vbCrLf
    strCode = strCode & "                             wsMitarbeiter.Cells(rngFound.Row, 3).Value" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetMitarbeiterName = """"" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Bereich des Mitarbeiters ermitteln (Innendienst/Außendienst)" & vbCrLf
    strCode = strCode & "Public Function GetMitarbeiterBereich(lngPersonalnummer As Long) As String" & vbCrLf
    strCode = strCode & "    Dim wsMitarbeiter As Worksheet" & vbCrLf
    strCode = strCode & "    Dim rngFound As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsMitarbeiter = ThisWorkbook.Sheets(""Mitarbeiter"")" & vbCrLf
    strCode = strCode & "    Set rngFound = wsMitarbeiter.Range(""A:A"").Find(lngPersonalnummer, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If Not rngFound Is Nothing Then" & vbCrLf
    strCode = strCode & "        GetMitarbeiterBereich = wsMitarbeiter.Cells(rngFound.Row, 6).Value" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetMitarbeiterBereich = """"" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Artikelname anhand ArtikelID ermitteln" & vbCrLf
    strCode = strCode & "Public Function GetArtikelName(intArtikelID As Integer) As String" & vbCrLf
    strCode = strCode & "    Dim wsSortiment As Worksheet" & vbCrLf
    strCode = strCode & "    Dim rngFound As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsSortiment = ThisWorkbook.Sheets(""Sortiment"")" & vbCrLf
    strCode = strCode & "    Set rngFound = wsSortiment.Range(""A:A"").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If Not rngFound Is Nothing Then" & vbCrLf
    strCode = strCode & "        GetArtikelName = wsSortiment.Cells(rngFound.Row, 2).Value" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetArtikelName = """"" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Neue Ausgabe hinzufügen" & vbCrLf
    strCode = strCode & "Public Sub AddAusgabe(dtDatum As Date, lngPersonalnummer As Long, _" & vbCrLf
    strCode = strCode & "                      intArtikelID As Integer, strGroesse As String, _" & vbCrLf
    strCode = strCode & "                      intMenge As Integer, strBemerkung As String)" & vbCrLf
    strCode = strCode & "    Dim wsAusgaben As Worksheet" & vbCrLf
    strCode = strCode & "    Dim lngNextRow As Long, lngNextID As Long" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsAusgaben = ThisWorkbook.Sheets(""Ausgaben"")" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Nächste freie Zeile finden" & vbCrLf
    strCode = strCode & "    lngNextRow = wsAusgaben.Cells(wsAusgaben.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    strCode = strCode & "    lngNextID = GetNextAusgabeID()" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Daten eintragen" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 1).Value = lngNextID" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 2).Value = dtDatum" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 3).Value = lngPersonalnummer" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 4).Formula = ""=IFERROR(VLOOKUP(C"" & lngNextRow & "",tblMitarbeiter,2,FALSE)&"""" """"&VLOOKUP(C"" & lngNextRow & "",tblMitarbeiter,3,FALSE),"""""""")""" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 5).Value = intArtikelID" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 6).Formula = ""=IFERROR(VLOOKUP(E"" & lngNextRow & "",tblSortiment,2,FALSE),"""""""")""" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 7).Value = strGroesse" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 8).Value = intMenge" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 9).Formula = ""=YEAR(B"" & lngNextRow & "")""" & vbCrLf
    strCode = strCode & "    wsAusgaben.Cells(lngNextRow, 10).Value = strBemerkung" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf & vbCrLf
    strCode = strCode & "' Nächste AusgabeID ermitteln" & vbCrLf
    strCode = strCode & "Public Function GetNextAusgabeID() As Long" & vbCrLf
    strCode = strCode & "    Dim wsAusgaben As Worksheet" & vbCrLf
    strCode = strCode & "    Dim lngMaxID As Long" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsAusgaben = ThisWorkbook.Sheets(""Ausgaben"")" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    On Error Resume Next" & vbCrLf
    strCode = strCode & "    lngMaxID = Application.WorksheetFunction.Max(wsAusgaben.Range(""A:A""))" & vbCrLf
    strCode = strCode & "    On Error GoTo 0" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    GetNextAusgabeID = lngMaxID + 1" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Standard-Anspruch für Artikel abrufen" & vbCrLf
    strCode = strCode & "Public Function GetStandardAnspruch(intArtikelID As Integer) As Integer" & vbCrLf
    strCode = strCode & "    Dim wsSortiment As Worksheet" & vbCrLf
    strCode = strCode & "    Dim rngFound As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsSortiment = ThisWorkbook.Sheets(""Sortiment"")" & vbCrLf
    strCode = strCode & "    Set rngFound = wsSortiment.Range(""A:A"").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If Not rngFound Is Nothing Then" & vbCrLf
    strCode = strCode & "        GetStandardAnspruch = CInt(wsSortiment.Cells(rngFound.Row, 3).Value)" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetStandardAnspruch = 0" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Zyklus-Jahre für Artikel abrufen" & vbCrLf
    strCode = strCode & "Public Function GetZyklusJahre(intArtikelID As Integer) As Integer" & vbCrLf
    strCode = strCode & "    Dim wsSortiment As Worksheet" & vbCrLf
    strCode = strCode & "    Dim rngFound As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsSortiment = ThisWorkbook.Sheets(""Sortiment"")" & vbCrLf
    strCode = strCode & "    Set rngFound = wsSortiment.Range(""A:A"").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If Not rngFound Is Nothing Then" & vbCrLf
    strCode = strCode & "        GetZyklusJahre = CInt(wsSortiment.Cells(rngFound.Row, 4).Value)" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetZyklusJahre = 1" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Zyklus-Typ für Artikel abrufen (Kalender/Rollierend)" & vbCrLf
    strCode = strCode & "Public Function GetZyklusTyp(intArtikelID As Integer) As String" & vbCrLf
    strCode = strCode & "    Dim wsSortiment As Worksheet" & vbCrLf
    strCode = strCode & "    Dim rngFound As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsSortiment = ThisWorkbook.Sheets(""Sortiment"")" & vbCrLf
    strCode = strCode & "    Set rngFound = wsSortiment.Range(""A:A"").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If Not rngFound Is Nothing Then" & vbCrLf
    strCode = strCode & "        GetZyklusTyp = wsSortiment.Cells(rngFound.Row, 5).Value" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetZyklusTyp = ""Kalender""" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf

    GetModDatenCode = strCode
End Function

' ============================================================================
' VBA-CODE: modBerechnung
' ============================================================================
Function GetModBerechnungCode()
    Dim strCode
    strCode = "Option Explicit" & vbCrLf & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf
    strCode = strCode & "' modBerechnung - Berechnungslogik für Ansprüche" & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf & vbCrLf
    strCode = strCode & "' Effektiven Anspruch berechnen (inkl. Sonderregel Innendienst)" & vbCrLf
    strCode = strCode & "Public Function BerechneEffektivenAnspruch(lngPersonalnummer As Long, _" & vbCrLf
    strCode = strCode & "                                           intArtikelID As Integer) As Integer" & vbCrLf
    strCode = strCode & "    Dim intStandard As Integer" & vbCrLf
    strCode = strCode & "    Dim strBereich As String" & vbCrLf
    strCode = strCode & "    Dim strArtikelName As String" & vbCrLf
    strCode = strCode & "    Dim wsConfig As Worksheet" & vbCrLf
    strCode = strCode & "    Dim intInnendienstAnspruch As Integer" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Standard-Anspruch aus Sortiment" & vbCrLf
    strCode = strCode & "    intStandard = modDaten.GetStandardAnspruch(intArtikelID)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Bereich und Artikelname ermitteln" & vbCrLf
    strCode = strCode & "    strBereich = modDaten.GetMitarbeiterBereich(lngPersonalnummer)" & vbCrLf
    strCode = strCode & "    strArtikelName = modDaten.GetArtikelName(intArtikelID)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Sonderregel: Innendienst bekommt nur 2 Hemden/Blusen" & vbCrLf
    strCode = strCode & "    If strBereich = ""Innendienst"" Then" & vbCrLf
    strCode = strCode & "        If strArtikelName = ""Hemd"" Or strArtikelName = ""Bluse"" Then" & vbCrLf
    strCode = strCode & "            ' Wert aus Config lesen" & vbCrLf
    strCode = strCode & "            Set wsConfig = ThisWorkbook.Sheets(""Config"")" & vbCrLf
    strCode = strCode & "            On Error Resume Next" & vbCrLf
    strCode = strCode & "            intInnendienstAnspruch = Application.WorksheetFunction.VLookup( _" & vbCrLf
    strCode = strCode & "                ""InnendienstHemdAnspruch"", wsConfig.Range(""A:B""), 2, False)" & vbCrLf
    strCode = strCode & "            On Error GoTo 0" & vbCrLf
    strCode = strCode & "            If intInnendienstAnspruch = 0 Then intInnendienstAnspruch = 2" & vbCrLf
    strCode = strCode & "            BerechneEffektivenAnspruch = intInnendienstAnspruch" & vbCrLf
    strCode = strCode & "            Exit Function" & vbCrLf
    strCode = strCode & "        End If" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Keine Sonderregel: Standard-Anspruch verwenden" & vbCrLf
    strCode = strCode & "    ' (Individuelle Abweichungen könnten hier noch ergänzt werden)" & vbCrLf
    strCode = strCode & "    BerechneEffektivenAnspruch = intStandard" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Restanspruch berechnen" & vbCrLf
    strCode = strCode & "Public Function BerechneRestanspruch(lngPersonalnummer As Long, _" & vbCrLf
    strCode = strCode & "                                     intArtikelID As Integer, _" & vbCrLf
    strCode = strCode & "                                     intJahr As Integer) As Integer" & vbCrLf
    strCode = strCode & "    Dim intEffektiverAnspruch As Integer" & vbCrLf
    strCode = strCode & "    Dim intAusgegeben As Integer" & vbCrLf
    strCode = strCode & "    Dim strZyklusTyp As String" & vbCrLf
    strCode = strCode & "    Dim intZyklusJahre As Integer" & vbCrLf
    strCode = strCode & "    Dim dtLetzteAusgabe As Date" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    intEffektiverAnspruch = BerechneEffektivenAnspruch(lngPersonalnummer, intArtikelID)" & vbCrLf
    strCode = strCode & "    strZyklusTyp = modDaten.GetZyklusTyp(intArtikelID)" & vbCrLf
    strCode = strCode & "    intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If strZyklusTyp = ""Kalender"" Then" & vbCrLf
    strCode = strCode & "        ' Jährlicher Anspruch: Ausgaben im Kalenderjahr zählen" & vbCrLf
    strCode = strCode & "        intAusgegeben = GetAusgabenImJahr(lngPersonalnummer, intArtikelID, intJahr)" & vbCrLf
    strCode = strCode & "        BerechneRestanspruch = intEffektiverAnspruch - intAusgegeben" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        ' Rollierender Anspruch: Prüfen ob seit letzter Ausgabe X Jahre vergangen" & vbCrLf
    strCode = strCode & "        dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)" & vbCrLf
    strCode = strCode & "        " & vbCrLf
    strCode = strCode & "        If dtLetzteAusgabe = 0 Then" & vbCrLf
    strCode = strCode & "            ' Noch nie ausgegeben -> voller Anspruch" & vbCrLf
    strCode = strCode & "            BerechneRestanspruch = intEffektiverAnspruch" & vbCrLf
    strCode = strCode & "        ElseIf intJahr - Year(dtLetzteAusgabe) >= intZyklusJahre Then" & vbCrLf
    strCode = strCode & "            ' Zyklus abgelaufen -> neuer Anspruch" & vbCrLf
    strCode = strCode & "            BerechneRestanspruch = intEffektiverAnspruch" & vbCrLf
    strCode = strCode & "        Else" & vbCrLf
    strCode = strCode & "            ' Noch im Zyklus -> kein Anspruch" & vbCrLf
    strCode = strCode & "            BerechneRestanspruch = 0" & vbCrLf
    strCode = strCode & "        End If" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Mindestens 0" & vbCrLf
    strCode = strCode & "    If BerechneRestanspruch < 0 Then BerechneRestanspruch = 0" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Ausgaben im Jahr zählen" & vbCrLf
    strCode = strCode & "Public Function GetAusgabenImJahr(lngPersonalnummer As Long, _" & vbCrLf
    strCode = strCode & "                                   intArtikelID As Integer, _" & vbCrLf
    strCode = strCode & "                                   intJahr As Integer) As Integer" & vbCrLf
    strCode = strCode & "    Dim wsAusgaben As Worksheet" & vbCrLf
    strCode = strCode & "    Dim dblSumme As Double" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsAusgaben = ThisWorkbook.Sheets(""Ausgaben"")" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    On Error Resume Next" & vbCrLf
    strCode = strCode & "    dblSumme = Application.WorksheetFunction.SumIfs( _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""H:H""), _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""C:C""), lngPersonalnummer, _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""E:E""), intArtikelID, _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""I:I""), intJahr)" & vbCrLf
    strCode = strCode & "    On Error GoTo 0" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    GetAusgabenImJahr = CInt(dblSumme)" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Letzte Ausgabe eines Artikels für einen Mitarbeiter" & vbCrLf
    strCode = strCode & "Public Function GetLetzteAusgabeDatum(lngPersonalnummer As Long, _" & vbCrLf
    strCode = strCode & "                                      intArtikelID As Integer) As Date" & vbCrLf
    strCode = strCode & "    Dim wsAusgaben As Worksheet" & vbCrLf
    strCode = strCode & "    Dim varMax As Variant" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsAusgaben = ThisWorkbook.Sheets(""Ausgaben"")" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    On Error Resume Next" & vbCrLf
    strCode = strCode & "    varMax = Application.WorksheetFunction.MaxIfs( _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""B:B""), _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""C:C""), lngPersonalnummer, _" & vbCrLf
    strCode = strCode & "        wsAusgaben.Range(""E:E""), intArtikelID)" & vbCrLf
    strCode = strCode & "    On Error GoTo 0" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    If IsError(varMax) Or IsEmpty(varMax) Or varMax = 0 Then" & vbCrLf
    strCode = strCode & "        GetLetzteAusgabeDatum = 0" & vbCrLf
    strCode = strCode & "    Else" & vbCrLf
    strCode = strCode & "        GetLetzteAusgabeDatum = CDate(varMax)" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Restanspruch berechnen und auf Blatt anzeigen" & vbCrLf
    strCode = strCode & "Public Sub BerechneUndZeigeRestanspruch(lngPersonalnummer As Long, intJahr As Integer)" & vbCrLf
    strCode = strCode & "    Dim wsRest As Worksheet, wsSortiment As Worksheet" & vbCrLf
    strCode = strCode & "    Dim lngRow As Long, intArtikelID As Integer" & vbCrLf
    strCode = strCode & "    Dim intStandard As Integer, intEffektiv As Integer" & vbCrLf
    strCode = strCode & "    Dim intAusgegeben As Integer, intRest As Integer" & vbCrLf
    strCode = strCode & "    Dim strStatus As String, strZyklusTyp As String" & vbCrLf
    strCode = strCode & "    Dim dtLetzteAusgabe As Date, intZyklusJahre As Integer" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsRest = ThisWorkbook.Sheets(""Restanspruch"")" & vbCrLf
    strCode = strCode & "    Set wsSortiment = ThisWorkbook.Sheets(""Sortiment"")" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Application.ScreenUpdating = False" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Alte Daten löschen (ab Zeile 8)" & vbCrLf
    strCode = strCode & "    wsRest.Range(""A8:F100"").ClearContents" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Durch alle aktiven Artikel iterieren" & vbCrLf
    strCode = strCode & "    lngRow = 8" & vbCrLf
    strCode = strCode & "    For Each rngArtikel In wsSortiment.ListObjects(""tblSortiment"").DataBodyRange.Rows" & vbCrLf
    strCode = strCode & "        If rngArtikel.Cells(1, 6).Value = ""Ja"" Then" & vbCrLf
    strCode = strCode & "            intArtikelID = CInt(rngArtikel.Cells(1, 1).Value)" & vbCrLf
    strCode = strCode & "            " & vbCrLf
    strCode = strCode & "            intStandard = modDaten.GetStandardAnspruch(intArtikelID)" & vbCrLf
    strCode = strCode & "            intEffektiv = BerechneEffektivenAnspruch(lngPersonalnummer, intArtikelID)" & vbCrLf
    strCode = strCode & "            intRest = BerechneRestanspruch(lngPersonalnummer, intArtikelID, intJahr)" & vbCrLf
    strCode = strCode & "            strZyklusTyp = modDaten.GetZyklusTyp(intArtikelID)" & vbCrLf
    strCode = strCode & "            intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)" & vbCrLf
    strCode = strCode & "            " & vbCrLf
    strCode = strCode & "            If strZyklusTyp = ""Kalender"" Then" & vbCrLf
    strCode = strCode & "                intAusgegeben = GetAusgabenImJahr(lngPersonalnummer, intArtikelID, intJahr)" & vbCrLf
    strCode = strCode & "                If intRest > 0 Then" & vbCrLf
    strCode = strCode & "                    strStatus = ""Verfügbar""" & vbCrLf
    strCode = strCode & "                Else" & vbCrLf
    strCode = strCode & "                    strStatus = ""Erschöpft""" & vbCrLf
    strCode = strCode & "                End If" & vbCrLf
    strCode = strCode & "            Else" & vbCrLf
    strCode = strCode & "                dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)" & vbCrLf
    strCode = strCode & "                intAusgegeben = 0 ' Bei rollierend nicht jahresbezogen" & vbCrLf
    strCode = strCode & "                If dtLetzteAusgabe = 0 Then" & vbCrLf
    strCode = strCode & "                    strStatus = ""Verfügbar (noch nie ausgegeben)""" & vbCrLf
    strCode = strCode & "                ElseIf intRest > 0 Then" & vbCrLf
    strCode = strCode & "                    strStatus = ""Verfügbar (letzte: "" & Year(dtLetzteAusgabe) & "")""" & vbCrLf
    strCode = strCode & "                Else" & vbCrLf
    strCode = strCode & "                    strStatus = ""Nächste: "" & (Year(dtLetzteAusgabe) + intZyklusJahre)" & vbCrLf
    strCode = strCode & "                End If" & vbCrLf
    strCode = strCode & "            End If" & vbCrLf
    strCode = strCode & "            " & vbCrLf
    strCode = strCode & "            wsRest.Cells(lngRow, 1).Value = modDaten.GetArtikelName(intArtikelID)" & vbCrLf
    strCode = strCode & "            wsRest.Cells(lngRow, 2).Value = intStandard" & vbCrLf
    strCode = strCode & "            wsRest.Cells(lngRow, 3).Value = intEffektiv" & vbCrLf
    strCode = strCode & "            wsRest.Cells(lngRow, 4).Value = intAusgegeben" & vbCrLf
    strCode = strCode & "            wsRest.Cells(lngRow, 5).Value = intRest" & vbCrLf
    strCode = strCode & "            wsRest.Cells(lngRow, 6).Value = strStatus" & vbCrLf
    strCode = strCode & "            " & vbCrLf
    strCode = strCode & "            lngRow = lngRow + 1" & vbCrLf
    strCode = strCode & "        End If" & vbCrLf
    strCode = strCode & "    Next rngArtikel" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Application.ScreenUpdating = True" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf & vbCrLf
    strCode = strCode & "' Übersicht aktualisieren" & vbCrLf
    strCode = strCode & "Public Sub AktualisiereUebersicht()" & vbCrLf
    strCode = strCode & "    Dim wsUebersicht As Worksheet, wsMitarbeiter As Worksheet" & vbCrLf
    strCode = strCode & "    Dim wsSortiment As Worksheet" & vbCrLf
    strCode = strCode & "    Dim lngRow As Long, lngCol As Long" & vbCrLf
    strCode = strCode & "    Dim rngMitarbeiter As Range, rngArtikel As Range" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Set wsUebersicht = ThisWorkbook.Sheets(""Uebersicht"")" & vbCrLf
    strCode = strCode & "    Set wsMitarbeiter = ThisWorkbook.Sheets(""Mitarbeiter"")" & vbCrLf
    strCode = strCode & "    Set wsSortiment = ThisWorkbook.Sheets(""Sortiment"")" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Application.ScreenUpdating = False" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Header und Daten löschen (ab Zeile 6)" & vbCrLf
    strCode = strCode & "    wsUebersicht.Range(""A6:Z1000"").ClearContents" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Mitarbeiter eintragen" & vbCrLf
    strCode = strCode & "    lngRow = 6" & vbCrLf
    strCode = strCode & "    For Each rngMitarbeiter In wsMitarbeiter.ListObjects(""tblMitarbeiter"").DataBodyRange.Rows" & vbCrLf
    strCode = strCode & "        If rngMitarbeiter.Cells(1, 5).Value = ""Ja"" Then" & vbCrLf
    strCode = strCode & "            wsUebersicht.Cells(lngRow, 1).Value = rngMitarbeiter.Cells(1, 1).Value" & vbCrLf
    strCode = strCode & "            wsUebersicht.Cells(lngRow, 2).Formula = ""=IFERROR(VLOOKUP(A"" & lngRow & "",tblMitarbeiter,2,FALSE)&"""" """"&VLOOKUP(A"" & lngRow & "",tblMitarbeiter,3,FALSE),"""""""")""" & vbCrLf
    strCode = strCode & "            " & vbCrLf
    strCode = strCode & "            lngCol = 3" & vbCrLf
    strCode = strCode & "            For Each rngArtikel In wsSortiment.ListObjects(""tblSortiment"").DataBodyRange.Rows" & vbCrLf
    strCode = strCode & "                If rngArtikel.Cells(1, 6).Value = ""Ja"" Then" & vbCrLf
    strCode = strCode & "                    wsUebersicht.Cells(5, lngCol).Value = rngArtikel.Cells(1, 2).Value" & vbCrLf
    strCode = strCode & "                    wsUebersicht.Cells(lngRow, lngCol).Formula = _" & vbCrLf
    strCode = strCode & "                        ""=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A"" & lngRow & _" & vbCrLf
    strCode = strCode & "                        "",tblAusgaben[ArtikelID],"" & rngArtikel.Cells(1, 1).Value & _" & vbCrLf
    strCode = strCode & "                        "",tblAusgaben[Kalenderjahr],$B$3)""" & vbCrLf
    strCode = strCode & "                    lngCol = lngCol + 1" & vbCrLf
    strCode = strCode & "                End If" & vbCrLf
    strCode = strCode & "            Next rngArtikel" & vbCrLf
    strCode = strCode & "            " & vbCrLf
    strCode = strCode & "            lngRow = lngRow + 1" & vbCrLf
    strCode = strCode & "        End If" & vbCrLf
    strCode = strCode & "    Next rngMitarbeiter" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Header formatieren" & vbCrLf
    strCode = strCode & "    With wsUebersicht.Range(""A5:G5"")" & vbCrLf
    strCode = strCode & "        .Font.Bold = True" & vbCrLf
    strCode = strCode & "        .Interior.Color = RGB(68, 114, 196)" & vbCrLf
    strCode = strCode & "        .Font.Color = RGB(255, 255, 255)" & vbCrLf
    strCode = strCode & "    End With" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    Application.ScreenUpdating = True" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf

    GetModBerechnungCode = strCode
End Function

' ============================================================================
' VBA-CODE: modHelfer
' ============================================================================
Function GetModHelferCode()
    Dim strCode
    strCode = "Option Explicit" & vbCrLf & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf
    strCode = strCode & "' modHelfer - Hilfsfunktionen" & vbCrLf
    strCode = strCode & "' ============================================================================" & vbCrLf & vbCrLf
    strCode = strCode & "' Ausgabe validieren" & vbCrLf
    strCode = strCode & "Public Function ValidateAusgabe(dtDatum As Date, lngPersonalnummer As Long, _" & vbCrLf
    strCode = strCode & "                                intArtikelID As Integer, intMenge As Integer) As Boolean" & vbCrLf
    strCode = strCode & "    ValidateAusgabe = True" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Datum nicht in Zukunft" & vbCrLf
    strCode = strCode & "    If dtDatum > Date Then" & vbCrLf
    strCode = strCode & "        MsgBox ""Datum darf nicht in der Zukunft liegen."", vbExclamation, APP_NAME" & vbCrLf
    strCode = strCode & "        ValidateAusgabe = False" & vbCrLf
    strCode = strCode & "        Exit Function" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Mitarbeiter existiert" & vbCrLf
    strCode = strCode & "    If modDaten.GetMitarbeiterName(lngPersonalnummer) = """" Then" & vbCrLf
    strCode = strCode & "        MsgBox ""Mitarbeiter nicht gefunden."", vbExclamation, APP_NAME" & vbCrLf
    strCode = strCode & "        ValidateAusgabe = False" & vbCrLf
    strCode = strCode & "        Exit Function" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Artikel existiert" & vbCrLf
    strCode = strCode & "    If modDaten.GetArtikelName(intArtikelID) = """" Then" & vbCrLf
    strCode = strCode & "        MsgBox ""Artikel nicht gefunden."", vbExclamation, APP_NAME" & vbCrLf
    strCode = strCode & "        ValidateAusgabe = False" & vbCrLf
    strCode = strCode & "        Exit Function" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "    " & vbCrLf
    strCode = strCode & "    ' Menge positiv" & vbCrLf
    strCode = strCode & "    If intMenge <= 0 Then" & vbCrLf
    strCode = strCode & "        MsgBox ""Menge muss größer als 0 sein."", vbExclamation, APP_NAME" & vbCrLf
    strCode = strCode & "        ValidateAusgabe = False" & vbCrLf
    strCode = strCode & "        Exit Function" & vbCrLf
    strCode = strCode & "    End If" & vbCrLf
    strCode = strCode & "End Function" & vbCrLf & vbCrLf
    strCode = strCode & "' Benannte Bereiche aktualisieren" & vbCrLf
    strCode = strCode & "Public Sub RefreshNamedRanges()" & vbCrLf
    strCode = strCode & "    ' Wird bei Bedarf implementiert" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf & vbCrLf
    strCode = strCode & "' Dropdowns aktualisieren" & vbCrLf
    strCode = strCode & "Public Sub AktualisiereDropdowns()" & vbCrLf
    strCode = strCode & "    ' Wird bei Bedarf implementiert" & vbCrLf
    strCode = strCode & "End Sub" & vbCrLf

    GetModHelferCode = strCode
End Function
