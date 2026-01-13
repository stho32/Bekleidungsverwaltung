# Excel VBA Konventionen und Best Practices

Diese Dokumentation definiert die Standards für die Excel-VBA-Entwicklung im Projekt Bekleidungsverwaltung.

## Inhaltsverzeichnis

1. [VBA Namenskonventionen](#vba-namenskonventionen)
2. [Tabellenblatt-Benennung](#tabellenblatt-benennung)
3. [Benannte Bereiche](#benannte-bereiche)
4. [Projektstruktur](#projektstruktur)
5. [Error Handling](#error-handling)
6. [Performance-Optimierung](#performance-optimierung)

---

## VBA Namenskonventionen

### Variablen-Benennung

Format: **[Scope] + [Datentyp] + [Name]**

#### Scope-Präfixe

| Präfix | Bedeutung | Beispiel |
|--------|-----------|----------|
| `m` | Modul-Level (Private) | `mstrCurrentUser` |
| `g` | Global (Public) | `gintMaxRows` |
| (ohne) | Lokal | `strName` |

#### Datentyp-Präfixe

| Typ | Präfix | Beispiel |
|-----|--------|----------|
| String | `str` | `strMitarbeiterName` |
| Integer | `int` | `intAnzahl` |
| Long | `lng` | `lngPersonalnummer` |
| Boolean | `bln` | `blnIstAktiv` |
| Date | `dt` | `dtAusgabeDatum` |
| Double | `dbl` | `dblProzent` |
| Currency | `cur` | `curBetrag` |
| Variant | `var` | `varDaten` |
| Array | `arr` | `arrMitarbeiter` |
| Range | `rng` | `rngAusgaben` |
| Worksheet | `ws` | `wsMitarbeiter` |
| Workbook | `wb` | `wbAktuell` |
| Object | `obj` | `objListBox` |

#### Beispiele

```vba
' Lokal
Dim strMitarbeiterName As String
Dim intAnzahl As Integer
Dim rngAusgaben As Range

' Modul-Level
Private mstrAktiverBenutzer As String
Private mintMaxZeilen As Integer

' Global
Public gstrAppVersion As String
Public gdtStartDatum As Date
```

### Konstanten-Benennung

Konstanten verwenden GROSSBUCHSTABEN mit Unterstrichen:

```vba
' Modul-Konstanten
Private Const MAX_MITARBEITER As Integer = 500
Private Const SORTIMENT_BLATT As String = "Sortiment"

' Globale Konstanten
Public Const APP_NAME As String = "Bekleidungsverwaltung"
Public Const VERSION As String = "1.0.0"
```

### Prozedur-Benennung

Format: **VerbSubstantiv** in PascalCase

```vba
' Subs
Public Sub LadeAusgaben()
Public Sub AktualisiereDaten()
Private Sub BerechneRestanspruch()

' Functions
Public Function GetMitarbeiterName(lngID As Long) As String
Public Function BerechneAnspruch(dtJahr As Date) As Integer
Private Function IstBerechtigt(lngMitarbeiterID As Long) As Boolean
```

### Formular-Steuerelemente

| Steuerelement | Präfix | Beispiel |
|---------------|--------|----------|
| CommandButton | `cmd` | `cmdSpeichern` |
| TextBox | `txt` | `txtPersonalnummer` |
| Label | `lbl` | `lblStatus` |
| ComboBox | `cbo` | `cboMitarbeiter` |
| ListBox | `lst` | `lstArtikel` |
| CheckBox | `chk` | `chkAktiv` |
| OptionButton | `opt` | `optJahr2025` |
| Frame | `fra` | `fraFilter` |

---

## Tabellenblatt-Benennung

### Regeln

- **Maximal 31 Zeichen**
- **Keine Leerzeichen** → Unterstriche verwenden
- **Keine Sonderzeichen**: `: \ / ? * [ ]`
- **Nicht "History"** (reserviert)
- **Beschreibend aber kurz**

### Konvention für dieses Projekt

| Blatt | Name | CodeName (VBA) |
|-------|------|----------------|
| Mitarbeiter-Stammdaten | `Mitarbeiter` | `wsMitarbeiter` |
| Sortiment/Artikel | `Sortiment` | `wsSortiment` |
| Ausgabe-Erfassung | `Ausgaben` | `wsAusgaben` |
| Jahresübersicht | `Uebersicht` | `wsUebersicht` |
| Restanspruch-Abfrage | `Restanspruch` | `wsRestanspruch` |
| Konfiguration | `Config` | `wsConfig` |

### VBA-Zugriff

```vba
' Empfohlen: Über CodeName (stabil bei Umbenennung)
wsMitarbeiter.Range("A1").Value = "Test"

' Alternativ: Über Blattname (anfällig bei Umbenennung)
ThisWorkbook.Sheets("Mitarbeiter").Range("A1").Value = "Test"
```

---

## Benannte Bereiche

### Namensregeln

- **Keine Leerzeichen** → Unterstriche oder CamelCase
- **Keine Zellreferenzen** wie A1, R1C1
- **Max. 255 Zeichen**
- **Beschreibend**: `Mitarbeiter_Liste` statt `Bereich1`

### Konvention für dieses Projekt

| Bereich | Name | Beschreibung |
|---------|------|--------------|
| Mitarbeiterliste | `tbl_Mitarbeiter` | Gesamte Mitarbeitertabelle |
| Mitarbeiter-Namen | `lst_Mitarbeiter` | Nur Namen für Dropdown |
| Artikelliste | `tbl_Sortiment` | Gesamte Sortimenttabelle |
| Artikel-Namen | `lst_Artikel` | Nur Artikelnamen für Dropdown |
| Ausgaben | `tbl_Ausgaben` | Alle Ausgabe-Einträge |

### Dynamische benannte Bereiche

Für wachsende Listen verwenden wir **INDEX mit COUNTA** (performanter als OFFSET):

```
=Mitarbeiter!$A$2:INDEX(Mitarbeiter!$A:$A,COUNTA(Mitarbeiter!$A:$A))
```

Alternative mit OFFSET (volatil, langsamer):

```
=OFFSET(Mitarbeiter!$A$2,0,0,COUNTA(Mitarbeiter!$A$2:$A$500),1)
```

### Excel-Tabellen (ListObjects)

Bevorzugt: Strukturierte Tabellen mit `Strg+T`:

```vba
' Zugriff auf Tabelle
Dim tbl As ListObject
Set tbl = wsMitarbeiter.ListObjects("tbl_Mitarbeiter")

' Neue Zeile hinzufügen
Dim newRow As ListRow
Set newRow = tbl.ListRows.Add
newRow.Range(1) = lngPersonalnummer
newRow.Range(2) = strName
```

---

## Projektstruktur

### VBA-Projekt-Aufbau

```
VBAProject (Bekleidungsverwaltung.xlsm)
├── Microsoft Excel Objekte
│   ├── DieseArbeitsmappe (ThisWorkbook)
│   ├── wsMitarbeiter (Mitarbeiter)
│   ├── wsSortiment (Sortiment)
│   ├── wsAusgaben (Ausgaben)
│   ├── wsUebersicht (Uebersicht)
│   └── wsRestanspruch (Restanspruch)
├── Formulare
│   └── frmAusgabeErfassung
├── Module
│   ├── modMain (Hauptlogik)
│   ├── modDaten (Datenzugriff)
│   ├── modBerechnung (Berechnungen)
│   └── modHelfer (Hilfsfunktionen)
└── Klassenmodule
    └── clsMitarbeiter (optional)
```

### Modul-Direktiven

Jedes Modul beginnt mit:

```vba
Option Explicit
Option Private Module  ' Falls nur intern genutzt
```

---

## Error Handling

### Standard-Struktur

```vba
Public Sub BeispielProzedur()
    On Error GoTo ErrorHandler

    ' === Initialisierung ===
    Application.ScreenUpdating = False

    ' === Hauptlogik ===
    ' ... Code hier ...

CleanExit:
    ' === Aufräumen ===
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' === Fehlerbehandlung ===
    MsgBox "Fehler in BeispielProzedur:" & vbCrLf & _
           "Nr: " & Err.Number & vbCrLf & _
           "Beschreibung: " & Err.Description, _
           vbCritical, APP_NAME
    Resume CleanExit
End Sub
```

### Benutzerdefinierte Fehler

```vba
' Im Modul-Kopf definieren
Public Const ERR_MITARBEITER_NICHT_GEFUNDEN As Long = vbObjectError + 1001
Public Const ERR_KONTINGENT_ERSCHOEPFT As Long = vbObjectError + 1002

' Fehler auslösen
If rngMitarbeiter Is Nothing Then
    Err.Raise ERR_MITARBEITER_NICHT_GEFUNDEN, _
              "GetMitarbeiter", _
              "Mitarbeiter mit ID " & lngID & " nicht gefunden."
End If
```

### Vermeiden

```vba
' NIEMALS so:
On Error Resume Next
' ... viel Code ...
' Fehler werden verschluckt!

' STATTDESSEN gezielt:
On Error Resume Next
Set rng = ws.Range("Test")  ' Könnte fehlschlagen
On Error GoTo 0  ' Fehlerbehandlung wieder aktivieren

If rng Is Nothing Then
    ' Behandlung
End If
```

---

## Performance-Optimierung

### Standard-Wrapper für Makros

```vba
Public Sub OptimierteAusfuehrung()
    ' === Performance-Einstellungen speichern ===
    Dim blnScreenUpdating As Boolean
    Dim lngCalculation As Long
    Dim blnEnableEvents As Boolean

    blnScreenUpdating = Application.ScreenUpdating
    lngCalculation = Application.Calculation
    blnEnableEvents = Application.EnableEvents

    On Error GoTo ErrorHandler

    ' === Performance-Modus aktivieren ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' === Hauptlogik ===
    ' ... Code hier ...

CleanExit:
    ' === Einstellungen wiederherstellen ===
    Application.ScreenUpdating = blnScreenUpdating
    Application.Calculation = lngCalculation
    Application.EnableEvents = blnEnableEvents
    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

### Best Practices

1. **Bereiche vollständig qualifizieren**
   ```vba
   ' Gut
   ThisWorkbook.Sheets("Daten").Range("A1")

   ' Schlecht (implizit ActiveSheet)
   Range("A1")
   ```

2. **Arrays statt Zellzugriffe**
   ```vba
   ' Schnell: Array
   Dim arrDaten As Variant
   arrDaten = rng.Value

   ' Langsam: Einzelne Zellen
   For i = 1 To 1000
       cells(i, 1).Value = i
   Next
   ```

3. **Value2 statt Value**
   ```vba
   ' Schneller (keine Datumskonvertierung)
   varWert = rng.Value2
   ```

4. **With-Blöcke nutzen**
   ```vba
   With wsAusgaben
       .Range("A1").Value = dtDatum
       .Range("B1").Value = strMitarbeiter
       .Range("C1").Value = strArtikel
   End With
   ```

---

## Quellen

- [VBA Naming Conventions - Software Solutions Online](https://software-solutions-online.com/vba-naming-conventions/)
- [VBA Development Best Practices - Spreadsheet1](https://www.spreadsheet1.com/vba-development-best-practices.html)
- [VBA Error Handling - Excel Macro Mastery](https://excelmacromastery.com/vba-error-handling/)
- [Dynamic Named Ranges - Microsoft Learn](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/create-dynamic-defined-range)
- [Excel Worksheet Naming - BetterSolutions](https://bettersolutions.com/excel/worksheets/naming.htm)
- [VBA On Error Best Practices - Automate Excel](https://www.automateexcel.com/vba/error-handling/)
