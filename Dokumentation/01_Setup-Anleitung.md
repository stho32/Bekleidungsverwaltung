# Setup-Anleitung: Bekleidungsverwaltung

Diese Anleitung führt Sie Schritt für Schritt durch die Einrichtung der Bekleidungsverwaltung.

---

## Inhaltsverzeichnis

1. [Voraussetzungen](#1-voraussetzungen)
2. [Excel-Datei erstellen](#2-excel-datei-erstellen)
3. [VBA-Code importieren (falls nötig)](#3-vba-code-importieren-falls-nötig)
4. [Erste Konfiguration](#4-erste-konfiguration)
5. [Stammdaten einrichten](#5-stammdaten-einrichten)
6. [Funktionstest](#6-funktionstest)
7. [Fehlerbehebung](#7-fehlerbehebung)

---

## 1. Voraussetzungen

### Systemanforderungen

| Komponente | Anforderung |
|------------|-------------|
| Betriebssystem | Windows 10/11 |
| Microsoft Excel | Version 2016 oder neuer |
| Makros | Müssen aktiviert sein |

### Vor der Installation prüfen

1. **Excel-Version prüfen:**
   - Excel öffnen → **Datei** → **Konto** → Version ablesen
   - Mindestens Excel 2016 erforderlich

2. **Makro-Einstellungen prüfen:**
   - Excel öffnen → **Datei** → **Optionen** → **Trust Center**
   - **Einstellungen für das Trust Center...** klicken
   - **Makroeinstellungen** auswählen
   - Empfohlen: "Alle Makros mit Benachrichtigung deaktivieren"

---

## 2. Excel-Datei erstellen

Sie haben zwei Möglichkeiten, die Excel-Datei zu erstellen:

### Option A: PowerShell-Script (empfohlen)

**Schritt 1:** Navigieren Sie zum Projektordner

```
c:\Projekte\Maike\
```

**Schritt 2:** Rechtsklick auf `Bekleidungsverwaltung_Setup.ps1`

**Schritt 3:** Wählen Sie **"Mit PowerShell ausführen"**

![PowerShell ausführen](images/powershell-run.png)

**Schritt 4:** Warten Sie, bis Excel sich öffnet und die Datei erstellt wird

**Schritt 5:** Die Datei wird automatisch gespeichert als:
```
c:\Projekte\Maike\Bekleidungsverwaltung.xlsm
```

### Option B: VBScript

**Schritt 1:** Doppelklick auf `Bekleidungsverwaltung_Setup.vbs`

**Schritt 2:** Excel öffnet sich automatisch

**Schritt 3:** Bei erfolgreicher Erstellung erscheint eine Bestätigungsmeldung

---

## 3. VBA-Code importieren (falls nötig)

> **Hinweis:** Dieser Schritt ist nur erforderlich, wenn der VBA-Code nicht automatisch eingefügt wurde. Sie erkennen dies daran, dass beim Drücken von **Alt + F8** keine Makros angezeigt werden.

### VBA-Projektzugriff aktivieren

**Schritt 1:** Excel öffnen

**Schritt 2:** Klicken Sie auf **Datei** → **Optionen**

![Optionen öffnen](images/excel-options.png)

**Schritt 3:** Wählen Sie **Trust Center** in der linken Spalte

**Schritt 4:** Klicken Sie auf **Einstellungen für das Trust Center...**

**Schritt 5:** Wählen Sie **Makroeinstellungen**

**Schritt 6:** Aktivieren Sie das Kontrollkästchen:
```
☑ Zugriff auf das VBA-Projektobjektmodell vertrauen
```

**Schritt 7:** Klicken Sie auf **OK** und starten Sie Excel neu

### VBA-Module manuell importieren

**Schritt 1:** Öffnen Sie `Bekleidungsverwaltung.xlsm`

**Schritt 2:** Drücken Sie **Alt + F11** (VBA-Editor öffnen)

**Schritt 3:** Im Projektfenster (links): Rechtsklick auf **VBAProject (Bekleidungsverwaltung.xlsm)**

**Schritt 4:** Wählen Sie **Datei importieren...**

**Schritt 5:** Navigieren Sie zum Ordner `VBA\` und importieren Sie nacheinander:

| Reihenfolge | Datei | Beschreibung |
|-------------|-------|--------------|
| 1 | `modMain.bas` | Hauptmodul mit Button-Handlern |
| 2 | `modDaten.bas` | Datenzugriffsschicht |
| 3 | `modBerechnung.bas` | Berechnungslogik |
| 4 | `modHelfer.bas` | Hilfsfunktionen |

**Schritt 6:** Speichern Sie die Datei (**Strg + S**)

**Schritt 7:** Schließen Sie den VBA-Editor (**Alt + Q**)

---

## 4. Erste Konfiguration

### Config-Blatt anpassen

**Schritt 1:** Öffnen Sie die Datei `Bekleidungsverwaltung.xlsm`

**Schritt 2:** Aktivieren Sie Makros wenn gefragt:

![Makros aktivieren](images/enable-macros.png)

**Schritt 3:** Wechseln Sie zum Blatt **Config**

**Schritt 4:** Passen Sie die Parameter an Ihre Bedürfnisse an:

| Parameter | Empfohlener Wert | Beschreibung |
|-----------|------------------|--------------|
| StartJahr | Aktuelles Jahr | Ab wann Daten erfasst werden |
| MaxZeilenAusgaben | 10000 | Maximum für Ausgabeneinträge |
| InnendienstHemdAnspruch | 2 | Hemden für Innendienst-MA |

**Schritt 5:** Speichern Sie die Datei

---

## 5. Stammdaten einrichten

### 5.1 Sortiment konfigurieren

**Schritt 1:** Wechseln Sie zum Blatt **Sortiment**

**Schritt 2:** Überprüfen Sie die vordefinierten Artikel:

| ArtikelID | Artikelname | Anspruch | Zyklus |
|-----------|-------------|----------|--------|
| 1 | Hemd | 4 | 1 Jahr |
| 2 | Bluse | 4 | 1 Jahr |
| 3 | Polo Shirt | 2 | 1 Jahr |
| 4 | Hoodie | 1 | 3 Jahre (rollierend) |
| 5 | Softshelljacke | 1 | 3 Jahre (rollierend) |

**Schritt 3:** Passen Sie die Anspruchsmengen bei Bedarf an (Spalte C)

**Schritt 4:** Fügen Sie bei Bedarf weitere Artikel hinzu:
- Neue Zeile in der Tabelle anlegen
- Eindeutige ArtikelID vergeben
- Alle Felder ausfüllen

### 5.2 Mitarbeiter erfassen

**Schritt 1:** Wechseln Sie zum Blatt **Mitarbeiter**

**Schritt 2:** Löschen Sie die Beispieldaten (Zeilen 2-4)

**Schritt 3:** Erfassen Sie Ihre Mitarbeiter mit folgenden Daten:

| Spalte | Inhalt | Beispiel |
|--------|--------|----------|
| A | Personalnummer | 1001 |
| B | Nachname | Müller |
| C | Vorname | Hans |
| D | Eintrittsdatum | 15.03.2020 |
| E | Aktiv | Ja |
| F | Bereich | Außendienst |
| G | Abteilung | Vertrieb |

**Wichtig für Bereich:**
- `Außendienst` = Standard-Ansprüche
- `Innendienst` = Reduzierter Hemd/Blusen-Anspruch (2 statt 4)

**Schritt 4:** Speichern Sie die Datei

---

## 6. Funktionstest

### Test 1: Makros verfügbar

**Schritt 1:** Drücken Sie **Alt + F8**

**Schritt 2:** Folgende Makros sollten erscheinen:
- `BtnNeueAusgabe_Click`
- `BtnUebersichtAktualisieren_Click`
- `BtnRestanspruchBerechnen_Click`
- `BtnAusgabenSortieren_Click`
- `InitializeApplication`

✅ **Erfolgreich:** Alle Makros sind sichtbar

❌ **Problem:** Keine Makros sichtbar → siehe [VBA-Code importieren](#3-vba-code-importieren-falls-nötig)

### Test 2: Neue Ausgabe erfassen

**Schritt 1:** Drücken Sie **Alt + F8**

**Schritt 2:** Wählen Sie `BtnNeueAusgabe_Click` und klicken Sie **Ausführen**

**Schritt 3:** Geben Sie Testdaten ein:
- Datum: Heutiges Datum
- Personalnummer: Eine vorhandene Nummer
- ArtikelID: 1 (Hemd)
- Größe: L
- Menge: 1

**Schritt 4:** Wechseln Sie zum Blatt **Ausgaben**

✅ **Erfolgreich:** Der Eintrag erscheint in der Tabelle

### Test 3: Restanspruch prüfen

**Schritt 1:** Wechseln Sie zum Blatt **Restanspruch**

**Schritt 2:** Geben Sie ein Jahr und eine Personalnummer ein

**Schritt 3:** Drücken Sie **Alt + F8** → `BtnRestanspruchBerechnen_Click` → **Ausführen**

✅ **Erfolgreich:** Die Restansprüche werden angezeigt

### Test 4: Innendienst-Sonderregel

**Schritt 1:** Erfassen Sie einen Mitarbeiter mit Bereich = `Innendienst`

**Schritt 2:** Prüfen Sie den Restanspruch für diesen Mitarbeiter

✅ **Erfolgreich:** Hemd/Bluse zeigt nur 2 als effektiven Anspruch (nicht 4)

---

## 7. Fehlerbehebung

### Problem: "Makros wurden deaktiviert"

**Lösung:**
1. Schließen Sie die Datei
2. Öffnen Sie die Datei erneut
3. Klicken Sie auf **Inhalt aktivieren** in der gelben Leiste

### Problem: "Kompilierungsfehler" beim Ausführen

**Lösung:**
1. **Alt + F11** drücken
2. Menü **Debuggen** → **Kompilieren von VBAProject**
3. Fehlermeldung beachten und korrigieren

### Problem: Dropdowns zeigen keine Werte

**Lösung:**
1. Prüfen Sie, ob Daten im Blatt **Mitarbeiter** bzw. **Sortiment** vorhanden sind
2. Prüfen Sie, ob die Tabellen korrekt als `tblMitarbeiter` und `tblSortiment` benannt sind

### Problem: Formeln zeigen #BEZUG!

**Lösung:**
1. Prüfen Sie, ob die Tabellennamen korrekt sind
2. Im VBA-Editor: **Extras** → **Verweise** prüfen (keine fehlenden Verweise)

### Problem: Script startet nicht (PowerShell)

**Lösung:**
1. PowerShell als Administrator öffnen
2. Ausführen: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
3. Script erneut starten

---

## Nächste Schritte

Nach erfolgreichem Setup:

1. 📖 Lesen Sie die [Benutzerdokumentation](02_Benutzerdokumentation.md)
2. 👥 Schulen Sie die Benutzer
3. 📊 Beginnen Sie mit der Datenerfassung

---

## Support

Bei weiteren Fragen wenden Sie sich an:
- IT-Abteilung
- Projektverantwortlicher: [Name einfügen]
