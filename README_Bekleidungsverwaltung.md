# Bekleidungsverwaltung - Setup-Anleitung

## Übersicht

Excel-VBA-Lösung zur Verwaltung von Unternehmensbekleidungs-Kontingenten für 100+ Mitarbeiter.

## Schnellstart

### Option 1: PowerShell-Script (empfohlen)

1. Rechtsklick auf `Bekleidungsverwaltung_Setup.ps1`
2. "Mit PowerShell ausführen" wählen
3. Excel öffnet sich automatisch und erstellt die Datei

**Hinweis:** Falls der VBA-Code nicht automatisch eingefügt wird, siehe "Manueller VBA-Import" unten.

### Option 2: VBScript

1. Doppelklick auf `Bekleidungsverwaltung_Setup.vbs`
2. Excel öffnet sich automatisch

## Voraussetzungen für VBA-Import

Für den automatischen VBA-Import muss folgende Option aktiviert sein:

1. Excel öffnen
2. **Datei** → **Optionen** → **Trust Center**
3. **Einstellungen für das Trust Center...** klicken
4. **Makroeinstellungen** wählen
5. ✅ **Zugriff auf das VBA-Projektobjektmodell vertrauen** aktivieren
6. **OK** klicken und Excel neu starten

## Manueller VBA-Import

Falls der automatische Import nicht funktioniert:

1. `Bekleidungsverwaltung.xlsm` öffnen
2. **Alt + F11** drücken (VBA-Editor öffnen)
3. Rechtsklick auf **VBAProject (Bekleidungsverwaltung.xlsm)**
4. **Datei importieren...** wählen
5. Nacheinander importieren:
   - `VBA/modMain.bas`
   - `VBA/modDaten.bas`
   - `VBA/modBerechnung.bas`
   - `VBA/modHelfer.bas`
6. Speichern und VBA-Editor schließen

## Verfügbare Makros

Nach dem Setup stehen folgende Makros zur Verfügung (Alt + F8):

| Makro | Beschreibung |
|-------|--------------|
| `BtnNeueAusgabe_Click` | Neue Bekleidungsausgabe erfassen |
| `BtnUebersichtAktualisieren_Click` | Jahresübersicht aktualisieren |
| `BtnRestanspruchBerechnen_Click` | Restanspruch für Mitarbeiter berechnen |
| `BtnAusgabenSortieren_Click` | Ausgaben nach Datum sortieren |
| `InitializeApplication` | Anwendung initialisieren |

## Tabellenblätter

| Blatt | Zweck |
|-------|-------|
| **Mitarbeiter** | Stammdaten (Personalnr., Name, Bereich, Aktiv-Status) |
| **Sortiment** | Artikel mit Anspruchsmengen und Zyklen |
| **Ausgaben** | Erfassung aller Bekleidungsausgaben |
| **Uebersicht** | Jahresauswertung pro Mitarbeiter |
| **Restanspruch** | Abfrage der verbleibenden Ansprüche |
| **Config** | Systemkonfiguration |

## Geschäftsregeln

### Jährliche Artikel (Kalender-Zyklus)
- Hemd: 4 Stück pro Jahr (Außendienst) / 2 Stück (Innendienst)
- Bluse: 4 Stück pro Jahr (Außendienst) / 2 Stück (Innendienst)
- Polo Shirt: 2 Stück pro Jahr

### Rollierende Artikel (3-Jahres-Zyklus)
- Hoodie: 1 Stück alle 3 Jahre
- Softshelljacke: 1 Stück alle 3 Jahre

**Sonderregel Innendienst:** Mitarbeiter mit Bereich "Innendienst" erhalten nur 2 Hemden/Blusen pro Jahr.

## Konfiguration

Im Blatt **Config** können folgende Parameter angepasst werden:

| Parameter | Standard | Beschreibung |
|-----------|----------|--------------|
| `StartJahr` | 2025 | Erstes Jahr für Datenerfassung |
| `MaxZeilenAusgaben` | 10000 | Performance-Limit |
| `InnendienstHemdAnspruch` | 2 | Hemd/Blusen-Anspruch für Innendienst |

## Dateien

```
Bekleidungsverwaltung.xlsm      # Hauptdatei (wird erstellt)
Bekleidungsverwaltung_Setup.ps1 # PowerShell Setup-Script
Bekleidungsverwaltung_Setup.vbs # VBScript Setup-Script
VBA/
  modMain.bas                   # Hauptmodul
  modDaten.bas                  # Datenzugriff
  modBerechnung.bas             # Berechnungslogik
  modHelfer.bas                 # Hilfsfunktionen
```

## Support

Bei Fragen oder Problemen wenden Sie sich an die IT-Abteilung.
