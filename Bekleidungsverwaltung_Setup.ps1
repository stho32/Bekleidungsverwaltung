# ============================================================================
# Bekleidungsverwaltung Setup Script (PowerShell)
# ============================================================================
# Dieses Script erstellt die komplette Excel-Lösung für die
# Bekleidungskontingent-Verwaltung.
#
# Ausführung: Rechtsklick -> Mit PowerShell ausführen
# oder in PowerShell: .\Bekleidungsverwaltung_Setup.ps1
# ============================================================================

$ErrorActionPreference = "Stop"

# Pfad zur Zieldatei
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$fileName = Join-Path $scriptPath "Bekleidungsverwaltung.xlsm"

Write-Host "Erstelle Bekleidungsverwaltung.xlsm..." -ForegroundColor Cyan

# Excel starten
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

try {
    # Neue Arbeitsmappe erstellen
    $workbook = $excel.Workbooks.Add()

    # ========================================================================
    # TABELLENBLÄTTER ERSTELLEN
    # ========================================================================
    $sheetNames = @("Mitarbeiter", "Sortiment", "Ausgaben", "Uebersicht", "Restanspruch", "Config")

    # Vorhandene Blätter entfernen (außer erstem)
    while ($workbook.Sheets.Count -gt 1) {
        $workbook.Sheets.Item($workbook.Sheets.Count).Delete()
    }

    # Erstes Blatt umbenennen
    $workbook.Sheets.Item(1).Name = $sheetNames[0]

    # Weitere Blätter hinzufügen
    for ($i = 1; $i -lt $sheetNames.Count; $i++) {
        $newSheet = $workbook.Sheets.Add([System.Reflection.Missing]::Value, $workbook.Sheets.Item($workbook.Sheets.Count))
        $newSheet.Name = $sheetNames[$i]
    }

    Write-Host "  Tabellenblätter erstellt" -ForegroundColor Green

    # ========================================================================
    # MITARBEITER-BLATT
    # ========================================================================
    $ws = $workbook.Sheets.Item("Mitarbeiter")

    # Überschriften
    $ws.Range("A1").Value2 = "Personalnummer"
    $ws.Range("B1").Value2 = "Nachname"
    $ws.Range("C1").Value2 = "Vorname"
    $ws.Range("D1").Value2 = "Eintrittsdatum"
    $ws.Range("E1").Value2 = "Aktiv"
    $ws.Range("F1").Value2 = "Bereich"
    $ws.Range("G1").Value2 = "Abteilung"

    # Formatierung
    $headerRange = $ws.Range("A1:G1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xC47244  # RGB(68, 114, 196) in BGR
    $headerRange.Font.Color = 0xFFFFFF

    # Beispieldaten
    $ws.Range("A2").Value2 = 1001
    $ws.Range("B2").Value2 = "Müller"
    $ws.Range("C2").Value2 = "Hans"
    $ws.Range("D2").Value2 = "15.03.2020"
    $ws.Range("E2").Value2 = "Ja"
    $ws.Range("F2").Value2 = "Außendienst"
    $ws.Range("G2").Value2 = "Vertrieb"

    $ws.Range("A3").Value2 = 1002
    $ws.Range("B3").Value2 = "Schmidt"
    $ws.Range("C3").Value2 = "Anna"
    $ws.Range("D3").Value2 = "01.07.2019"
    $ws.Range("E3").Value2 = "Ja"
    $ws.Range("F3").Value2 = "Innendienst"
    $ws.Range("G3").Value2 = "Buchhaltung"

    $ws.Range("A4").Value2 = 1003
    $ws.Range("B4").Value2 = "Weber"
    $ws.Range("C4").Value2 = "Thomas"
    $ws.Range("D4").Value2 = "10.01.2021"
    $ws.Range("E4").Value2 = "Ja"
    $ws.Range("F4").Value2 = "Außendienst"
    $ws.Range("G4").Value2 = "Technik"

    # Datenvalidierung für Aktiv
    $validation = $ws.Range("E2:E1000").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "Ja,Nein")

    # Datenvalidierung für Bereich
    $validation = $ws.Range("F2:F1000").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "Außendienst,Innendienst")

    # Als Tabelle formatieren
    $listObj = $ws.ListObjects.Add(1, $ws.Range("A1:G4"), $null, 1)
    $listObj.Name = "tblMitarbeiter"

    $ws.Columns.Item("A:G").AutoFit()

    Write-Host "  Mitarbeiter-Blatt erstellt" -ForegroundColor Green

    # ========================================================================
    # SORTIMENT-BLATT
    # ========================================================================
    $ws = $workbook.Sheets.Item("Sortiment")

    # Überschriften
    $ws.Range("A1").Value2 = "ArtikelID"
    $ws.Range("B1").Value2 = "Artikelname"
    $ws.Range("C1").Value2 = "AnspruchMenge"
    $ws.Range("D1").Value2 = "ZyklusJahre"
    $ws.Range("E1").Value2 = "ZyklusTyp"
    $ws.Range("F1").Value2 = "Aktiv"
    $ws.Range("G1").Value2 = "Groessen"

    $headerRange = $ws.Range("A1:G1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xC47244
    $headerRange.Font.Color = 0xFFFFFF

    # Artikel
    $articles = @(
        @(1, "Hemd", 4, 1, "Kalender", "Ja", "S,M,L,XL,XXL"),
        @(2, "Bluse", 4, 1, "Kalender", "Ja", "XS,S,M,L,XL"),
        @(3, "Polo Shirt", 2, 1, "Kalender", "Ja", "S,M,L,XL,XXL"),
        @(4, "Hoodie", 1, 3, "Rollierend", "Ja", "S,M,L,XL,XXL"),
        @(5, "Softshelljacke", 1, 3, "Rollierend", "Ja", "S,M,L,XL,XXL")
    )

    for ($i = 0; $i -lt $articles.Count; $i++) {
        $row = $i + 2
        $ws.Range("A$row").Value2 = $articles[$i][0]
        $ws.Range("B$row").Value2 = $articles[$i][1]
        $ws.Range("C$row").Value2 = $articles[$i][2]
        $ws.Range("D$row").Value2 = $articles[$i][3]
        $ws.Range("E$row").Value2 = $articles[$i][4]
        $ws.Range("F$row").Value2 = $articles[$i][5]
        $ws.Range("G$row").Value2 = $articles[$i][6]
    }

    # Datenvalidierung
    $validation = $ws.Range("E2:E100").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "Kalender,Rollierend")

    $validation = $ws.Range("F2:F100").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "Ja,Nein")

    # Als Tabelle formatieren
    $listObj = $ws.ListObjects.Add(1, $ws.Range("A1:G6"), $null, 1)
    $listObj.Name = "tblSortiment"

    $ws.Columns.Item("A:G").AutoFit()

    Write-Host "  Sortiment-Blatt erstellt" -ForegroundColor Green

    # ========================================================================
    # AUSGABEN-BLATT
    # ========================================================================
    $ws = $workbook.Sheets.Item("Ausgaben")

    # Überschriften
    $ws.Range("A1").Value2 = "AusgabeID"
    $ws.Range("B1").Value2 = "Datum"
    $ws.Range("C1").Value2 = "Personalnummer"
    $ws.Range("D1").Value2 = "MitarbeiterName"
    $ws.Range("E1").Value2 = "ArtikelID"
    $ws.Range("F1").Value2 = "Artikelname"
    $ws.Range("G1").Value2 = "Groesse"
    $ws.Range("H1").Value2 = "Menge"
    $ws.Range("I1").Value2 = "Kalenderjahr"
    $ws.Range("J1").Value2 = "Bemerkung"

    $headerRange = $ws.Range("A1:J1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xC47244
    $headerRange.Font.Color = 0xFFFFFF

    # Beispieldaten
    $ws.Range("A2").Value2 = 1
    $ws.Range("B2").Value2 = "15.01.2025"
    $ws.Range("C2").Value2 = 1001
    $ws.Range("D2").Formula = "=IFERROR(VLOOKUP(C2,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(C2,tblMitarbeiter,3,FALSE),"""")"
    $ws.Range("E2").Value2 = 1
    $ws.Range("F2").Formula = "=IFERROR(VLOOKUP(E2,tblSortiment,2,FALSE),"""")"
    $ws.Range("G2").Value2 = "L"
    $ws.Range("H2").Value2 = 2
    $ws.Range("I2").Formula = "=YEAR(B2)"

    $ws.Range("A3").Value2 = 2
    $ws.Range("B3").Value2 = "15.01.2025"
    $ws.Range("C3").Value2 = 1002
    $ws.Range("D3").Formula = "=IFERROR(VLOOKUP(C3,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(C3,tblMitarbeiter,3,FALSE),"""")"
    $ws.Range("E3").Value2 = 2
    $ws.Range("F3").Formula = "=IFERROR(VLOOKUP(E3,tblSortiment,2,FALSE),"""")"
    $ws.Range("G3").Value2 = "M"
    $ws.Range("H3").Value2 = 1
    $ws.Range("I3").Formula = "=YEAR(B3)"

    # Datumsformat
    $ws.Columns.Item("B").NumberFormat = "DD.MM.YYYY"

    # Als Tabelle formatieren
    $listObj = $ws.ListObjects.Add(1, $ws.Range("A1:J3"), $null, 1)
    $listObj.Name = "tblAusgaben"

    $ws.Columns.Item("A:J").AutoFit()

    Write-Host "  Ausgaben-Blatt erstellt" -ForegroundColor Green

    # ========================================================================
    # ÜBERSICHT-BLATT
    # ========================================================================
    $ws = $workbook.Sheets.Item("Uebersicht")

    $ws.Range("A1").Value2 = "Ausgabenübersicht nach Jahr"
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Size = 14

    $ws.Range("A3").Value2 = "Jahr:"
    $ws.Range("B3").Value2 = 2025

    $validation = $ws.Range("B3").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "2024,2025,2026,2027,2028,2029,2030")

    # Header
    $ws.Range("A5").Value2 = "Personalnummer"
    $ws.Range("B5").Value2 = "Name"
    $ws.Range("C5").Value2 = "Hemd"
    $ws.Range("D5").Value2 = "Bluse"
    $ws.Range("E5").Value2 = "Polo Shirt"
    $ws.Range("F5").Value2 = "Hoodie"
    $ws.Range("G5").Value2 = "Softshelljacke"

    $headerRange = $ws.Range("A5:G5")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xC47244
    $headerRange.Font.Color = 0xFFFFFF

    # Datenzeilen mit Formeln
    for ($row = 6; $row -le 8; $row++) {
        $pnr = 1000 + $row - 5
        $ws.Range("A$row").Value2 = $pnr
        $ws.Range("B$row").Formula = "=IFERROR(VLOOKUP(A$row,tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(A$row,tblMitarbeiter,3,FALSE),"""")"
        $ws.Range("C$row").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],`$A$row,tblAusgaben[ArtikelID],1,tblAusgaben[Kalenderjahr],`$B`$3)"
        $ws.Range("D$row").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],`$A$row,tblAusgaben[ArtikelID],2,tblAusgaben[Kalenderjahr],`$B`$3)"
        $ws.Range("E$row").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],`$A$row,tblAusgaben[ArtikelID],3,tblAusgaben[Kalenderjahr],`$B`$3)"
        $ws.Range("F$row").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],`$A$row,tblAusgaben[ArtikelID],4,tblAusgaben[Kalenderjahr],`$B`$3)"
        $ws.Range("G$row").Formula = "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],`$A$row,tblAusgaben[ArtikelID],5,tblAusgaben[Kalenderjahr],`$B`$3)"
    }

    $ws.Range("A10").Value2 = "Hinweis: Klicken Sie auf 'Übersicht aktualisieren' um alle Mitarbeiter anzuzeigen."
    $ws.Range("A10").Font.Italic = $true

    $ws.Columns.Item("A:G").AutoFit()

    Write-Host "  Übersicht-Blatt erstellt" -ForegroundColor Green

    # ========================================================================
    # RESTANSPRUCH-BLATT
    # ========================================================================
    $ws = $workbook.Sheets.Item("Restanspruch")

    $ws.Range("A1").Value2 = "Restanspruch-Abfrage"
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Size = 14

    $ws.Range("A3").Value2 = "Jahr:"
    $ws.Range("B3").Value2 = 2025

    $ws.Range("A4").Value2 = "Mitarbeiter (Personalnr.):"
    $ws.Range("B4").Value2 = 1001

    $inputRange = $ws.Range("B3:B4")
    $inputRange.Interior.Color = 0xC8FFFF  # Hellgelb
    $inputRange.Borders.LineStyle = 1

    $validation = $ws.Range("B3").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "2024,2025,2026,2027,2028,2029,2030")

    # Ergebnis-Header
    $ws.Range("A7").Value2 = "Artikel"
    $ws.Range("B7").Value2 = "Standard"
    $ws.Range("C7").Value2 = "Effektiv"
    $ws.Range("D7").Value2 = "Ausgegeben"
    $ws.Range("E7").Value2 = "Rest"
    $ws.Range("F7").Value2 = "Status"

    $headerRange = $ws.Range("A7:F7")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xC47244
    $headerRange.Font.Color = 0xFFFFFF

    # Beispieldaten
    $resultData = @(
        @("Hemd", 4, 4, 2, 2, "Verfügbar"),
        @("Bluse", 4, 4, 0, 4, "Verfügbar"),
        @("Polo Shirt", 2, 2, 0, 2, "Verfügbar"),
        @("Hoodie", 1, 1, 0, 1, "Verfügbar (3-Jahres-Zyklus)"),
        @("Softshelljacke", 1, 1, 0, 1, "Verfügbar (3-Jahres-Zyklus)")
    )

    for ($i = 0; $i -lt $resultData.Count; $i++) {
        $row = $i + 8
        $ws.Range("A$row").Value2 = $resultData[$i][0]
        $ws.Range("B$row").Value2 = $resultData[$i][1]
        $ws.Range("C$row").Value2 = $resultData[$i][2]
        $ws.Range("D$row").Value2 = $resultData[$i][3]
        $ws.Range("E$row").Value2 = $resultData[$i][4]
        $ws.Range("F$row").Value2 = $resultData[$i][5]
    }

    $ws.Range("A14").Value2 = "Hinweis: Klicken Sie auf 'Berechnen' um den aktuellen Restanspruch anzuzeigen."
    $ws.Range("A14").Font.Italic = $true

    $ws.Columns.Item("A:F").AutoFit()

    Write-Host "  Restanspruch-Blatt erstellt" -ForegroundColor Green

    # ========================================================================
    # CONFIG-BLATT
    # ========================================================================
    $ws = $workbook.Sheets.Item("Config")

    $ws.Range("A1").Value2 = "Systemkonfiguration"
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Size = 14

    $ws.Range("A3").Value2 = "Parameter"
    $ws.Range("B3").Value2 = "Wert"
    $ws.Range("C3").Value2 = "Beschreibung"

    $headerRange = $ws.Range("A3:C3")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xC47244
    $headerRange.Font.Color = 0xFFFFFF

    $configData = @(
        @("StartJahr", 2025, "Erstes Jahr für Datenerfassung"),
        @("MaxZeilenAusgaben", 10000, "Maximale Anzahl Ausgabe-Einträge"),
        @("InnendienstHemdAnspruch", 2, "Reduzierter Hemd/Blusen-Anspruch für Innendienst"),
        @("AppVersion", "1.0.0", "Version der Anwendung")
    )

    for ($i = 0; $i -lt $configData.Count; $i++) {
        $row = $i + 4
        $ws.Range("A$row").Value2 = $configData[$i][0]
        $ws.Range("B$row").Value2 = $configData[$i][1]
        $ws.Range("C$row").Value2 = $configData[$i][2]
    }

    # Als Tabelle formatieren
    $listObj = $ws.ListObjects.Add(1, $ws.Range("A3:C7"), $null, 1)
    $listObj.Name = "tblConfig"

    $ws.Columns.Item("A:C").AutoFit()

    Write-Host "  Config-Blatt erstellt" -ForegroundColor Green

    # ========================================================================
    # VBA-CODE EINFÜGEN
    # ========================================================================
    Write-Host "  Füge VBA-Code ein..." -ForegroundColor Yellow

    try {
        $vbProject = $workbook.VBProject

        # modMain
        $module = $vbProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        $module.Name = "modMain"
        $module.CodeModule.AddFromString((Get-ModMainCode))

        # modDaten
        $module = $vbProject.VBComponents.Add(1)
        $module.Name = "modDaten"
        $module.CodeModule.AddFromString((Get-ModDatenCode))

        # modBerechnung
        $module = $vbProject.VBComponents.Add(1)
        $module.Name = "modBerechnung"
        $module.CodeModule.AddFromString((Get-ModBerechnungCode))

        # modHelfer
        $module = $vbProject.VBComponents.Add(1)
        $module.Name = "modHelfer"
        $module.CodeModule.AddFromString((Get-ModHelferCode))

        Write-Host "  VBA-Code eingefügt" -ForegroundColor Green
    }
    catch {
        Write-Host "  WARNUNG: VBA-Code konnte nicht eingefügt werden." -ForegroundColor Red
        Write-Host "  Bitte aktivieren Sie 'Zugriff auf das VBA-Projektobjektmodell vertrauen'" -ForegroundColor Red
        Write-Host "  unter Datei > Optionen > Trust Center > Einstellungen > Makroeinstellungen" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Die Excel-Datei wird ohne VBA-Code gespeichert." -ForegroundColor Yellow
        Write-Host "  Sie können den VBA-Code manuell aus den .bas-Dateien importieren." -ForegroundColor Yellow
    }

    # ========================================================================
    # SPEICHERN
    # ========================================================================
    $workbook.SaveAs($fileName, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled

    Write-Host ""
    Write-Host "Erfolgreich erstellt: $fileName" -ForegroundColor Green
    Write-Host ""

}
catch {
    Write-Host "Fehler: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    $excel.DisplayAlerts = $true
}

# ============================================================================
# VBA-CODE FUNKTIONEN
# ============================================================================

function Get-ModMainCode {
    return @'
Option Explicit

' ============================================================================
' modMain - Hauptmodul für Bekleidungsverwaltung
' ============================================================================

Public Const APP_NAME As String = "Bekleidungsverwaltung"
Public Const APP_VERSION As String = "1.0.0"

' Anwendung initialisieren
Public Sub InitializeApplication()
    On Error GoTo ErrorHandler

    Call modHelfer.AktualisiereDropdowns
    Call modHelfer.RefreshNamedRanges

    MsgBox "Anwendung initialisiert.", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler in InitializeApplication: " & Err.Description, vbCritical, APP_NAME
End Sub

' Neue Ausgabe hinzufügen
Public Sub BtnNeueAusgabe_Click()
    On Error GoTo ErrorHandler

    Dim strDatum As String, strPersonalnr As String
    Dim strArtikelID As String, strGroesse As String
    Dim strMenge As String, strBemerkung As String

    strDatum = InputBox("Datum (TT.MM.JJJJ):", APP_NAME, Format(Date, "DD.MM.YYYY"))
    If strDatum = "" Then Exit Sub

    strPersonalnr = InputBox("Personalnummer:", APP_NAME)
    If strPersonalnr = "" Then Exit Sub

    strArtikelID = InputBox("ArtikelID (1=Hemd, 2=Bluse, 3=Polo, 4=Hoodie, 5=Softshell):", APP_NAME)
    If strArtikelID = "" Then Exit Sub

    strGroesse = InputBox("Größe (S, M, L, XL, XXL):", APP_NAME, "L")
    If strGroesse = "" Then Exit Sub

    strMenge = InputBox("Menge:", APP_NAME, "1")
    If strMenge = "" Then Exit Sub

    strBemerkung = InputBox("Bemerkung (optional):", APP_NAME)

    Call modDaten.AddAusgabe(CDate(strDatum), CLng(strPersonalnr), _
                             CInt(strArtikelID), strGroesse, _
                             CInt(strMenge), strBemerkung)

    MsgBox "Ausgabe erfolgreich hinzugefügt!", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical, APP_NAME
End Sub

' Übersicht aktualisieren
Public Sub BtnUebersichtAktualisieren_Click()
    On Error GoTo ErrorHandler

    Call modBerechnung.AktualisiereUebersicht

    MsgBox "Übersicht aktualisiert!", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical, APP_NAME
End Sub

' Restanspruch berechnen
Public Sub BtnRestanspruchBerechnen_Click()
    On Error GoTo ErrorHandler

    Dim wsRest As Worksheet
    Dim intJahr As Integer, lngPersonalnr As Long

    Set wsRest = ThisWorkbook.Sheets("Restanspruch")
    intJahr = CInt(wsRest.Range("B3").Value)
    lngPersonalnr = CLng(wsRest.Range("B4").Value)

    Call modBerechnung.BerechneUndZeigeRestanspruch(lngPersonalnr, intJahr)
    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical, APP_NAME
End Sub
'@
}

function Get-ModDatenCode {
    return @'
Option Explicit

' ============================================================================
' modDaten - Datenzugriffsschicht
' ============================================================================

' Mitarbeitername anhand Personalnummer ermitteln
Public Function GetMitarbeiterName(lngPersonalnummer As Long) As String
    Dim wsMitarbeiter As Worksheet
    Dim rngFound As Range

    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set rngFound = wsMitarbeiter.Range("A:A").Find(lngPersonalnummer, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetMitarbeiterName = wsMitarbeiter.Cells(rngFound.Row, 2).Value & " " & _
                             wsMitarbeiter.Cells(rngFound.Row, 3).Value
    Else
        GetMitarbeiterName = ""
    End If
End Function

' Bereich des Mitarbeiters ermitteln
Public Function GetMitarbeiterBereich(lngPersonalnummer As Long) As String
    Dim wsMitarbeiter As Worksheet
    Dim rngFound As Range

    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set rngFound = wsMitarbeiter.Range("A:A").Find(lngPersonalnummer, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetMitarbeiterBereich = wsMitarbeiter.Cells(rngFound.Row, 6).Value
    Else
        GetMitarbeiterBereich = ""
    End If
End Function

' Artikelname anhand ArtikelID ermitteln
Public Function GetArtikelName(intArtikelID As Integer) As String
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetArtikelName = wsSortiment.Cells(rngFound.Row, 2).Value
    Else
        GetArtikelName = ""
    End If
End Function

' Neue Ausgabe hinzufügen
Public Sub AddAusgabe(dtDatum As Date, lngPersonalnummer As Long, _
                      intArtikelID As Integer, strGroesse As String, _
                      intMenge As Integer, strBemerkung As String)
    Dim wsAusgaben As Worksheet
    Dim lngNextRow As Long, lngNextID As Long

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")

    lngNextRow = wsAusgaben.Cells(wsAusgaben.Rows.Count, 1).End(xlUp).Row + 1
    lngNextID = GetNextAusgabeID()

    wsAusgaben.Cells(lngNextRow, 1).Value = lngNextID
    wsAusgaben.Cells(lngNextRow, 2).Value = dtDatum
    wsAusgaben.Cells(lngNextRow, 3).Value = lngPersonalnummer
    wsAusgaben.Cells(lngNextRow, 4).Formula = "=IFERROR(VLOOKUP(C" & lngNextRow & ",tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(C" & lngNextRow & ",tblMitarbeiter,3,FALSE),"""")"
    wsAusgaben.Cells(lngNextRow, 5).Value = intArtikelID
    wsAusgaben.Cells(lngNextRow, 6).Formula = "=IFERROR(VLOOKUP(E" & lngNextRow & ",tblSortiment,2,FALSE),"""")"
    wsAusgaben.Cells(lngNextRow, 7).Value = strGroesse
    wsAusgaben.Cells(lngNextRow, 8).Value = intMenge
    wsAusgaben.Cells(lngNextRow, 9).Formula = "=YEAR(B" & lngNextRow & ")"
    wsAusgaben.Cells(lngNextRow, 10).Value = strBemerkung
End Sub

' Nächste AusgabeID ermitteln
Public Function GetNextAusgabeID() As Long
    Dim wsAusgaben As Worksheet
    Dim lngMaxID As Long

    Set wsAusgaben = ThisWorkbook.Sheets("Ausgaben")

    On Error Resume Next
    lngMaxID = Application.WorksheetFunction.Max(wsAusgaben.Range("A:A"))
    On Error GoTo 0

    GetNextAusgabeID = lngMaxID + 1
End Function

' Standard-Anspruch für Artikel abrufen
Public Function GetStandardAnspruch(intArtikelID As Integer) As Integer
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetStandardAnspruch = CInt(wsSortiment.Cells(rngFound.Row, 3).Value)
    Else
        GetStandardAnspruch = 0
    End If
End Function

' Zyklus-Jahre für Artikel abrufen
Public Function GetZyklusJahre(intArtikelID As Integer) As Integer
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetZyklusJahre = CInt(wsSortiment.Cells(rngFound.Row, 4).Value)
    Else
        GetZyklusJahre = 1
    End If
End Function

' Zyklus-Typ für Artikel abrufen
Public Function GetZyklusTyp(intArtikelID As Integer) As String
    Dim wsSortiment As Worksheet
    Dim rngFound As Range

    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")
    Set rngFound = wsSortiment.Range("A:A").Find(intArtikelID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        GetZyklusTyp = wsSortiment.Cells(rngFound.Row, 5).Value
    Else
        GetZyklusTyp = "Kalender"
    End If
End Function
'@
}

function Get-ModBerechnungCode {
    return @'
Option Explicit

' ============================================================================
' modBerechnung - Berechnungslogik für Ansprüche
' ============================================================================

' Effektiven Anspruch berechnen (inkl. Sonderregel Innendienst)
Public Function BerechneEffektivenAnspruch(lngPersonalnummer As Long, _
                                           intArtikelID As Integer) As Integer
    Dim intStandard As Integer
    Dim strBereich As String
    Dim strArtikelName As String
    Dim wsConfig As Worksheet
    Dim intInnendienstAnspruch As Integer

    intStandard = modDaten.GetStandardAnspruch(intArtikelID)
    strBereich = modDaten.GetMitarbeiterBereich(lngPersonalnummer)
    strArtikelName = modDaten.GetArtikelName(intArtikelID)

    ' Sonderregel: Innendienst bekommt nur 2 Hemden/Blusen
    If strBereich = "Innendienst" Then
        If strArtikelName = "Hemd" Or strArtikelName = "Bluse" Then
            Set wsConfig = ThisWorkbook.Sheets("Config")
            On Error Resume Next
            intInnendienstAnspruch = Application.WorksheetFunction.VLookup( _
                "InnendienstHemdAnspruch", wsConfig.Range("A:B"), 2, False)
            On Error GoTo 0
            If intInnendienstAnspruch = 0 Then intInnendienstAnspruch = 2
            BerechneEffektivenAnspruch = intInnendienstAnspruch
            Exit Function
        End If
    End If

    BerechneEffektivenAnspruch = intStandard
End Function

' Restanspruch berechnen
Public Function BerechneRestanspruch(lngPersonalnummer As Long, _
                                     intArtikelID As Integer, _
                                     intJahr As Integer) As Integer
    Dim intEffektiverAnspruch As Integer
    Dim intAusgegeben As Integer
    Dim strZyklusTyp As String
    Dim intZyklusJahre As Integer
    Dim dtLetzteAusgabe As Date

    intEffektiverAnspruch = BerechneEffektivenAnspruch(lngPersonalnummer, intArtikelID)
    strZyklusTyp = modDaten.GetZyklusTyp(intArtikelID)
    intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)

    If strZyklusTyp = "Kalender" Then
        intAusgegeben = GetAusgabenImJahr(lngPersonalnummer, intArtikelID, intJahr)
        BerechneRestanspruch = intEffektiverAnspruch - intAusgegeben
    Else
        dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)

        If dtLetzteAusgabe = 0 Then
            BerechneRestanspruch = intEffektiverAnspruch
        ElseIf intJahr - Year(dtLetzteAusgabe) >= intZyklusJahre Then
            BerechneRestanspruch = intEffektiverAnspruch
        Else
            BerechneRestanspruch = 0
        End If
    End If

    If BerechneRestanspruch < 0 Then BerechneRestanspruch = 0
End Function

' Ausgaben im Jahr zählen
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

' Letzte Ausgabe ermitteln
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

' Restanspruch berechnen und auf Blatt anzeigen
Public Sub BerechneUndZeigeRestanspruch(lngPersonalnummer As Long, intJahr As Integer)
    Dim wsRest As Worksheet, wsSortiment As Worksheet
    Dim lngRow As Long, intArtikelID As Integer
    Dim intStandard As Integer, intEffektiv As Integer
    Dim intAusgegeben As Integer, intRest As Integer
    Dim strStatus As String, strZyklusTyp As String
    Dim dtLetzteAusgabe As Date, intZyklusJahre As Integer
    Dim rngArtikel As Range

    Set wsRest = ThisWorkbook.Sheets("Restanspruch")
    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")

    Application.ScreenUpdating = False

    wsRest.Range("A8:F100").ClearContents

    lngRow = 8
    For Each rngArtikel In wsSortiment.ListObjects("tblSortiment").DataBodyRange.Rows
        If rngArtikel.Cells(1, 6).Value = "Ja" Then
            intArtikelID = CInt(rngArtikel.Cells(1, 1).Value)

            intStandard = modDaten.GetStandardAnspruch(intArtikelID)
            intEffektiv = BerechneEffektivenAnspruch(lngPersonalnummer, intArtikelID)
            intRest = BerechneRestanspruch(lngPersonalnummer, intArtikelID, intJahr)
            strZyklusTyp = modDaten.GetZyklusTyp(intArtikelID)
            intZyklusJahre = modDaten.GetZyklusJahre(intArtikelID)

            If strZyklusTyp = "Kalender" Then
                intAusgegeben = GetAusgabenImJahr(lngPersonalnummer, intArtikelID, intJahr)
                If intRest > 0 Then
                    strStatus = "Verfügbar"
                Else
                    strStatus = "Erschöpft"
                End If
            Else
                dtLetzteAusgabe = GetLetzteAusgabeDatum(lngPersonalnummer, intArtikelID)
                intAusgegeben = 0
                If dtLetzteAusgabe = 0 Then
                    strStatus = "Verfügbar (noch nie ausgegeben)"
                ElseIf intRest > 0 Then
                    strStatus = "Verfügbar (letzte: " & Year(dtLetzteAusgabe) & ")"
                Else
                    strStatus = "Nächste: " & (Year(dtLetzteAusgabe) + intZyklusJahre)
                End If
            End If

            wsRest.Cells(lngRow, 1).Value = modDaten.GetArtikelName(intArtikelID)
            wsRest.Cells(lngRow, 2).Value = intStandard
            wsRest.Cells(lngRow, 3).Value = intEffektiv
            wsRest.Cells(lngRow, 4).Value = intAusgegeben
            wsRest.Cells(lngRow, 5).Value = intRest
            wsRest.Cells(lngRow, 6).Value = strStatus

            lngRow = lngRow + 1
        End If
    Next rngArtikel

    Application.ScreenUpdating = True
End Sub

' Übersicht aktualisieren
Public Sub AktualisiereUebersicht()
    Dim wsUebersicht As Worksheet, wsMitarbeiter As Worksheet
    Dim wsSortiment As Worksheet
    Dim lngRow As Long, lngCol As Long
    Dim rngMitarbeiter As Range, rngArtikel As Range

    Set wsUebersicht = ThisWorkbook.Sheets("Uebersicht")
    Set wsMitarbeiter = ThisWorkbook.Sheets("Mitarbeiter")
    Set wsSortiment = ThisWorkbook.Sheets("Sortiment")

    Application.ScreenUpdating = False

    wsUebersicht.Range("A6:Z1000").ClearContents

    lngRow = 6
    For Each rngMitarbeiter In wsMitarbeiter.ListObjects("tblMitarbeiter").DataBodyRange.Rows
        If rngMitarbeiter.Cells(1, 5).Value = "Ja" Then
            wsUebersicht.Cells(lngRow, 1).Value = rngMitarbeiter.Cells(1, 1).Value
            wsUebersicht.Cells(lngRow, 2).Formula = "=IFERROR(VLOOKUP(A" & lngRow & ",tblMitarbeiter,2,FALSE)&"" ""&VLOOKUP(A" & lngRow & ",tblMitarbeiter,3,FALSE),"""")"

            lngCol = 3
            For Each rngArtikel In wsSortiment.ListObjects("tblSortiment").DataBodyRange.Rows
                If rngArtikel.Cells(1, 6).Value = "Ja" Then
                    wsUebersicht.Cells(5, lngCol).Value = rngArtikel.Cells(1, 2).Value
                    wsUebersicht.Cells(lngRow, lngCol).Formula = _
                        "=SUMIFS(tblAusgaben[Menge],tblAusgaben[Personalnummer],$A" & lngRow & _
                        ",tblAusgaben[ArtikelID]," & rngArtikel.Cells(1, 1).Value & _
                        ",tblAusgaben[Kalenderjahr],$B$3)"
                    lngCol = lngCol + 1
                End If
            Next rngArtikel

            lngRow = lngRow + 1
        End If
    Next rngMitarbeiter

    With wsUebersicht.Range("A5:G5")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    Application.ScreenUpdating = True
End Sub
'@
}

function Get-ModHelferCode {
    return @'
Option Explicit

' ============================================================================
' modHelfer - Hilfsfunktionen
' ============================================================================

' Ausgabe validieren
Public Function ValidateAusgabe(dtDatum As Date, lngPersonalnummer As Long, _
                                intArtikelID As Integer, intMenge As Integer) As Boolean
    ValidateAusgabe = True

    If dtDatum > Date Then
        MsgBox "Datum darf nicht in der Zukunft liegen.", vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    If modDaten.GetMitarbeiterName(lngPersonalnummer) = "" Then
        MsgBox "Mitarbeiter nicht gefunden.", vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    If modDaten.GetArtikelName(intArtikelID) = "" Then
        MsgBox "Artikel nicht gefunden.", vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If

    If intMenge <= 0 Then
        MsgBox "Menge muss größer als 0 sein.", vbExclamation, APP_NAME
        ValidateAusgabe = False
        Exit Function
    End If
End Function

' Benannte Bereiche aktualisieren
Public Sub RefreshNamedRanges()
    ' Wird bei Bedarf implementiert
End Sub

' Dropdowns aktualisieren
Public Sub AktualisiereDropdowns()
    ' Wird bei Bedarf implementiert
End Sub
'@
}

Write-Host ""
Write-Host "Setup abgeschlossen!" -ForegroundColor Cyan
Write-Host ""
Write-Host "Nächste Schritte:" -ForegroundColor Yellow
Write-Host "1. Öffnen Sie die Datei: $fileName" -ForegroundColor White
Write-Host "2. Aktivieren Sie Makros wenn gefragt" -ForegroundColor White
Write-Host "3. Drücken Sie Alt+F8 um die Makros zu sehen" -ForegroundColor White
Write-Host ""
Write-Host "Verfügbare Makros:" -ForegroundColor Yellow
Write-Host "  - BtnNeueAusgabe_Click (Neue Ausgabe erfassen)" -ForegroundColor White
Write-Host "  - BtnUebersichtAktualisieren_Click (Übersicht aktualisieren)" -ForegroundColor White
Write-Host "  - BtnRestanspruchBerechnen_Click (Restanspruch berechnen)" -ForegroundColor White
Write-Host ""

Read-Host "Drücken Sie Enter zum Beenden"
