# Bekleidungsverwaltung

Excel-VBA-Lösung zur Verwaltung von Unternehmensbekleidungs-Kontingenten.

## Übersicht

Diese Anwendung ermöglicht die transparente Verwaltung von Bekleidungsansprüchen für 100+ Mitarbeiter mit:

- **Jährlichen Kontingenten** (Hemden, Blusen, Polo Shirts)
- **Rollierenden 3-Jahres-Kontingenten** (Hoodie, Softshelljacke)
- **Sonderregeln** für Innendienst-Mitarbeiter
- **Ausgabe-Erfassung** mit Größen und Mengen
- **Restanspruchs-Berechnung** pro Mitarbeiter und Jahr

## Schnellstart

### 1. Excel-Datei erstellen

```powershell
# PowerShell-Script ausführen (Rechtsklick → Mit PowerShell ausführen)
.\Bekleidungsverwaltung_Setup.ps1
```

### 2. Makros aktivieren

Beim Öffnen der erstellten `Bekleidungsverwaltung.xlsm` auf **"Inhalt aktivieren"** klicken.

### 3. Loslegen

| Aktion | Tastenkürzel |
|--------|--------------|
| Makro ausführen | Alt + F8 |
| Neue Ausgabe | `BtnNeueAusgabe_Click` |
| Restanspruch prüfen | `BtnRestanspruchBerechnen_Click` |

## Projektstruktur

```
Bekleidungsverwaltung/
├── Anforderungen/
│   └── R00001-Excel-Bekleidungskontingent-Verwaltung.md   # Anforderungsdokumentation
├── App-Architecture/
│   ├── Datenmodell.md                                      # Tabellenstruktur
│   └── Excel-VBA-Conventions.md                            # Coding-Standards
├── Dokumentation/
│   ├── 01_Setup-Anleitung.md                               # Installationsanleitung
│   └── 02_Benutzerdokumentation.md                         # Anwendungsfälle
├── VBA/
│   ├── modMain.bas                                         # Hauptmodul
│   ├── modDaten.bas                                        # Datenzugriff
│   ├── modBerechnung.bas                                   # Berechnungslogik
│   └── modHelfer.bas                                       # Hilfsfunktionen
├── Bekleidungsverwaltung_Setup.ps1                         # Setup-Script (PowerShell)
├── Bekleidungsverwaltung_Setup.vbs                         # Setup-Script (VBScript)
└── README.md                                               # Diese Datei
```

## Tabellenblätter

| Blatt | Zweck |
|-------|-------|
| **Mitarbeiter** | Stammdaten mit Personalnummer, Name, Bereich |
| **Sortiment** | Artikel mit Anspruchsmengen und Zyklen |
| **Ausgaben** | Erfassung aller Bekleidungsausgaben |
| **Uebersicht** | Jahresauswertung pro Mitarbeiter |
| **Restanspruch** | Abfrage verbleibender Ansprüche |
| **Config** | Systemkonfiguration |

## Geschäftsregeln

### Standard-Kontingente

| Artikel | Anspruch | Zyklus |
|---------|----------|--------|
| Hemd | 4 pro Jahr | Jährlich |
| Bluse | 4 pro Jahr | Jährlich |
| Polo Shirt | 2 pro Jahr | Jährlich |
| Hoodie | 1 Stück | Alle 3 Jahre |
| Softshelljacke | 1 Stück | Alle 3 Jahre |

### Sonderregel Innendienst

Mitarbeiter mit Bereich = "Innendienst" erhalten nur **2 Hemden/Blusen** pro Jahr (statt 4).

## Dokumentation

| Dokument | Beschreibung |
|----------|--------------|
| [Setup-Anleitung](Dokumentation/01_Setup-Anleitung.md) | Schritt-für-Schritt Installation |
| [Benutzerdokumentation](Dokumentation/02_Benutzerdokumentation.md) | 10 Anwendungsfälle erklärt |
| [Anforderungen](Anforderungen/R00001-Excel-Bekleidungskontingent-Verwaltung.md) | Fachliche Spezifikation |
| [Datenmodell](App-Architecture/Datenmodell.md) | Technische Tabellenstruktur |
| [VBA-Conventions](App-Architecture/Excel-VBA-Conventions.md) | Coding-Standards |

## Systemanforderungen

- Microsoft Excel 2016 oder neuer
- Windows 10/11
- Makros müssen aktiviert sein

## Lizenz

Internes Projekt - Alle Rechte vorbehalten.
