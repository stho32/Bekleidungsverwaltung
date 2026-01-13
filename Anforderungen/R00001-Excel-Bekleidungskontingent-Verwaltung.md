# R00001-Excel-Bekleidungskontingent-Verwaltung

## Zusammenfassung
Excel-basierte Lösung zur Verwaltung von Unternehmensbekleidungs-Kontingenten für 100+ Mitarbeiter mit jährlichen (Hemden, Polo Shirts) und rollierenden 3-Jahres-Ansprüchen (Hoodie, Softshelljacke), inklusive Ausgabe-Erfassung und Restanspruchs-Berechnung.

## Geschäftlicher Nutzen
Transparente Verwaltung der Bekleidungsansprüche aller Mitarbeiter, Vermeidung von Über- oder Unterversorgung, nachvollziehbare Dokumentation der Ausgaben.

## Funktionale Anforderungen

### Stammdatenverwaltung
- Mitarbeiterliste mit Personalnummer, Name, Eintrittsdatum auf separatem Blatt
- Konfigurierbares Sortiment mit Artikelbezeichnung, Anspruchsmenge und Zyklus (1 Jahr / 3 Jahre rollierend)
- Anspruchsmengen müssen ohne Formeländerung anpassbar sein

### Ausgabe-Erfassung
- Erfassungsblatt mit Spalten: Datum, Mitarbeiter (Dropdown), Artikel (Dropdown), Menge
- Validierung: Nur gültige Mitarbeiter und Artikel auswählbar
- Button "Neue Ausgabe hinzufügen" für komfortable Eingabe

### Auswertungen
- Übersichtsblatt: Ausgegebene Mengen pro Mitarbeiter gruppiert nach Kalenderjahr
- Restanspruchs-Rechner: Eingabe von Jahr und Mitarbeiter → Anzeige verbleibender Ansprüche pro Artikel
- Bei 3-Jahres-Artikeln: Berechnung ab letzter Ausgabe (rollierend)

### VBA-Funktionalität
- Makro zur automatischen Sortierung/Filterung der Ausgabeliste
- Makro zur Aktualisierung der Übersicht
- Eingabe-Formular für neue Ausgaben (optional)

## Nicht-funktionale Anforderungen
- **Performance**: Flüssige Bedienung bei 100+ Mitarbeitern und mehreren tausend Ausgabe-Einträgen
- **Benutzerfreundlichkeit**: Dropdown-Auswahlen statt Freitexteingabe, klare Beschriftungen
- **Wartbarkeit**: Sortiment und Ansprüche in separater Konfigurationstabelle, nicht in Formeln hartcodiert

## Akzeptanzkriterien
- [ ] Mitarbeiter können mit Personalnummer und Name erfasst werden
- [ ] Sortiment (Artikel + Anspruch) kann ohne Formeländerung erweitert/geändert werden
- [ ] Ausgaben können mit Datum, Mitarbeiter, Artikel und Menge erfasst werden
- [ ] Übersicht zeigt pro Mitarbeiter die Ausgaben gruppiert nach Kalenderjahr
- [ ] Restanspruch für einen Mitarbeiter in einem bestimmten Jahr ist abfragbar
- [ ] Bei 3-Jahres-Artikeln wird korrekt geprüft, ob seit letzter Ausgabe 3 Jahre vergangen sind
- [ ] Datei funktioniert mit aktivierten Makros in Excel 2016+

## Betroffene Verzeichnisstruktur
- Zieldatei: `./Bekleidungsverwaltung.xlsm` (Excel mit Makros)
- Optional Dokumentation: `./Anforderungen/R00001-Excel-Bekleidungskontingent-Verwaltung.md`

## Technische Überlegungen
- **Tabellenblätter**:
  1. `Mitarbeiter` - Stammdaten
  2. `Sortiment` - Artikel und Anspruchsregeln
  3. `Ausgaben` - Erfassungsliste
  4. `Übersicht` - Pivot-ähnliche Auswertung
  5. `Restanspruch` - Abfrage-Interface
- **Benannte Bereiche**: Für Mitarbeiterliste und Sortiment (dynamisch erweiterbar)
- **SUMMENPRODUKT/SUMMEWENNS**: Für Aggregationen nach Jahr und Mitarbeiter
- **MAXWENNS**: Für Ermittlung der letzten Ausgabe eines 3-Jahres-Artikels
- **VBA-Module**: Separate Module für Eingabe-Logik und Aktualisierung

## Abhängigkeiten
- Microsoft Excel 2016 oder neuer mit Makro-Unterstützung
- Aktivierte Makros beim Öffnen der Datei

## Geklärte Anforderungen

### Größenerfassung
**Entscheidung:** Ja, Größen werden pro Ausgabe erfasst.
- Größen (z.B. S, M, L, XL, XXL) werden bei jeder Ausgabe mit dokumentiert
- Verfügbare Größen werden pro Artikel im Sortiment hinterlegt

### Individuelle Ansprüche
**Entscheidung:** Es gibt einen generellen Standard, aber einzelne Mitarbeiter können Abweichungen haben.
- Standard-Ansprüche werden im Sortiment definiert
- Individuelle Abweichungen (z.B. +2 oder -1) werden pro Mitarbeiter und Artikel hinterlegt
- **Sonderregel Innendienst:** Mitarbeiter im Innendienst erhalten nur 2 Hemden/Blusen (statt Standard)

### Passwortschutz
**Entscheidung:** Nein, kein Passwortschutz für Konfigurationsblätter.

### Ausgeschiedene Mitarbeiter
**Entscheidung:** Aktiv-Flag verwenden.
- Mitarbeiter werden nicht gelöscht, sondern als "inaktiv" markiert
- Historische Ausgabedaten bleiben erhalten
- Inaktive Mitarbeiter erscheinen nicht mehr in Auswahl-Dropdowns

## Manuelle Vorbereitungstätigkeiten
- Mitarbeiterliste mit Personalnummern zusammenstellen
- Aktuelles Sortiment mit korrekten Anspruchsmengen definieren
- Entscheidung über Startjahr des 3-Jahres-Zyklus (z.B. 2025 als erstes Jahr)

## Manuelle Nachbereitungstätigkeiten
- Import der Mitarbeiter-Stammdaten in die Excel-Datei
- Test mit realen Beispieldaten
- Schulung der Teammitglieder zur Nutzung der Datei
- Festlegung, wer die Datei pflegen darf (Berechtigungen)

## Missing-Docs
- Keine bestehende Dokumentation zu Excel-Lösungen im Projekt vorhanden
- Unklar, ob es Vorgaben für VBA-Coding-Standards gibt
