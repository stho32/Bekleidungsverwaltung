# Benutzerdokumentation: Bekleidungsverwaltung

Diese Dokumentation erklärt die tägliche Nutzung der Bekleidungsverwaltung anhand praktischer Anwendungsfälle.

---

## Inhaltsverzeichnis

1. [Übersicht der Benutzeroberfläche](#1-übersicht-der-benutzeroberfläche)
2. [Anwendungsfälle](#2-anwendungsfälle)
   - [Fall 1: Neue Bekleidungsausgabe erfassen](#fall-1-neue-bekleidungsausgabe-erfassen)
   - [Fall 2: Restanspruch eines Mitarbeiters prüfen](#fall-2-restanspruch-eines-mitarbeiters-prüfen)
   - [Fall 3: Jahresübersicht anzeigen](#fall-3-jahresübersicht-anzeigen)
   - [Fall 4: Neuen Mitarbeiter anlegen](#fall-4-neuen-mitarbeiter-anlegen)
   - [Fall 5: Mitarbeiter deaktivieren (Austritt)](#fall-5-mitarbeiter-deaktivieren-austritt)
   - [Fall 6: Neuen Artikel zum Sortiment hinzufügen](#fall-6-neuen-artikel-zum-sortiment-hinzufügen)
   - [Fall 7: Anspruchsmengen anpassen](#fall-7-anspruchsmengen-anpassen)
   - [Fall 8: Ausgaben nach Datum sortieren](#fall-8-ausgaben-nach-datum-sortieren)
   - [Fall 9: Innendienst-Mitarbeiter korrekt anlegen](#fall-9-innendienst-mitarbeiter-korrekt-anlegen)
   - [Fall 10: 3-Jahres-Artikel verstehen](#fall-10-3-jahres-artikel-verstehen)
3. [Tipps und Best Practices](#3-tipps-und-best-practices)
4. [Häufige Fragen (FAQ)](#4-häufige-fragen-faq)

---

## 1. Übersicht der Benutzeroberfläche

### Tabellenblätter

| Blatt | Symbol | Zweck | Wer nutzt es? |
|-------|--------|-------|---------------|
| **Mitarbeiter** | 👤 | Stammdaten aller Mitarbeiter | Administrator |
| **Sortiment** | 👕 | Verfügbare Bekleidungsartikel | Administrator |
| **Ausgaben** | 📋 | Liste aller Bekleidungsausgaben | Alle Benutzer |
| **Uebersicht** | 📊 | Jahresauswertung pro Mitarbeiter | Alle Benutzer |
| **Restanspruch** | 🔍 | Abfrage verbleibender Ansprüche | Alle Benutzer |
| **Config** | ⚙️ | Systemeinstellungen | Administrator |

### Verfügbare Makros (Alt + F8)

| Makro | Tastenkürzel | Beschreibung |
|-------|--------------|--------------|
| `BtnNeueAusgabe_Click` | - | Neue Ausgabe erfassen |
| `BtnRestanspruchBerechnen_Click` | - | Restanspruch berechnen |
| `BtnUebersichtAktualisieren_Click` | - | Übersicht aktualisieren |
| `BtnAusgabenSortieren_Click` | - | Ausgaben sortieren |

---

## 2. Anwendungsfälle

---

### Fall 1: Neue Bekleidungsausgabe erfassen

**Szenario:** Ein Mitarbeiter erhält neue Arbeitskleidung und dies muss dokumentiert werden.

#### Schritt-für-Schritt

**Schritt 1:** Öffnen Sie die Datei `Bekleidungsverwaltung.xlsm`

**Schritt 2:** Drücken Sie **Alt + F8** um die Makro-Liste zu öffnen

**Schritt 3:** Wählen Sie `BtnNeueAusgabe_Click` und klicken Sie **Ausführen**

**Schritt 4:** Füllen Sie die Dialogfelder aus:

| Feld | Eingabe | Beispiel |
|------|---------|----------|
| Datum | TT.MM.JJJJ | 15.01.2025 |
| Personalnummer | Nummer des Mitarbeiters | 1001 |
| ArtikelID | Nummer aus Sortiment | 1 (Hemd) |
| Größe | XS, S, M, L, XL, XXL | L |
| Menge | Anzahl | 2 |
| Bemerkung | Optional | Erstausstattung |

**Schritt 5:** Bestätigen Sie mit **OK**

**Ergebnis:** Die Ausgabe erscheint im Blatt **Ausgaben** als neue Zeile.

#### Alternative: Direkteingabe im Ausgaben-Blatt

1. Wechseln Sie zum Blatt **Ausgaben**
2. Gehen Sie zur letzten Zeile der Tabelle
3. Geben Sie die Daten manuell ein
4. Die Formeln für MitarbeiterName und Artikelname werden automatisch berechnet

---

### Fall 2: Restanspruch eines Mitarbeiters prüfen

**Szenario:** Vor einer Ausgabe soll geprüft werden, wie viel Bekleidung ein Mitarbeiter noch erhalten kann.

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Restanspruch**

**Schritt 2:** Geben Sie die Abfrageparameter ein:

| Feld | Zelle | Eingabe |
|------|-------|---------|
| Jahr | B3 | 2025 |
| Personalnummer | B4 | 1001 |

**Schritt 3:** Drücken Sie **Alt + F8**

**Schritt 4:** Wählen Sie `BtnRestanspruchBerechnen_Click` und klicken Sie **Ausführen**

**Ergebnis:** Die Tabelle zeigt für jeden Artikel:

| Spalte | Bedeutung |
|--------|-----------|
| Artikel | Name des Bekleidungsstücks |
| Standard | Anspruch laut Sortiment |
| Effektiv | Tatsächlicher Anspruch (nach Sonderregeln) |
| Ausgegeben | Bereits erhaltene Menge im Jahr |
| Rest | Noch verfügbarer Anspruch |
| Status | Verfügbar / Erschöpft / Nächste Berechtigung |

#### Ergebnisse interpretieren

**Für jährliche Artikel (Hemd, Bluse, Polo):**
- ✅ "Verfügbar" = Mitarbeiter kann noch Artikel erhalten
- ❌ "Erschöpft" = Anspruch für dieses Jahr aufgebraucht

**Für 3-Jahres-Artikel (Hoodie, Softshelljacke):**
- ✅ "Verfügbar (letzte: 2022)" = Zyklus abgelaufen, neuer Anspruch
- ❌ "Nächste: 2028" = Nächste Berechtigung erst in Zukunft

---

### Fall 3: Jahresübersicht anzeigen

**Szenario:** Eine Übersicht aller Ausgaben für ein bestimmtes Jahr wird benötigt.

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Uebersicht**

**Schritt 2:** Wählen Sie das gewünschte Jahr in Zelle **B3**

**Schritt 3:** Drücken Sie **Alt + F8**

**Schritt 4:** Wählen Sie `BtnUebersichtAktualisieren_Click` und klicken Sie **Ausführen**

**Ergebnis:** Die Matrix zeigt für jeden aktiven Mitarbeiter die Anzahl der ausgegebenen Artikel.

#### Übersicht lesen

```
                    | Hemd | Bluse | Polo | Hoodie | Softshell |
--------------------|------|-------|------|--------|-----------|
Müller Hans         |  2   |   0   |  1   |   0    |     0     |
Schmidt Anna        |  0   |   1   |  0   |   1    |     0     |
Weber Thomas        |  4   |   0   |  2   |   0    |     1     |
```

- Zahl > 0 = Mitarbeiter hat diese Menge erhalten
- Zahl = 0 = Keine Ausgabe in diesem Jahr

---

### Fall 4: Neuen Mitarbeiter anlegen

**Szenario:** Ein neuer Mitarbeiter tritt ins Unternehmen ein.

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Mitarbeiter**

**Schritt 2:** Klicken Sie in die erste leere Zeile der Tabelle

**Schritt 3:** Geben Sie die Daten ein:

| Spalte | Feld | Pflicht | Beispiel |
|--------|------|---------|----------|
| A | Personalnummer | ✅ | 1004 |
| B | Nachname | ✅ | Meyer |
| C | Vorname | ✅ | Lisa |
| D | Eintrittsdatum | ✅ | 01.02.2025 |
| E | Aktiv | ✅ | Ja |
| F | Bereich | ✅ | Außendienst |
| G | Abteilung | ❌ | Marketing |

**Wichtig:**
- Personalnummer muss eindeutig sein
- Bereich bestimmt die Ansprüche (siehe [Fall 9](#fall-9-innendienst-mitarbeiter-korrekt-anlegen))

**Schritt 4:** Speichern Sie die Datei

**Ergebnis:** Der Mitarbeiter erscheint ab sofort in den Dropdown-Listen.

---

### Fall 5: Mitarbeiter deaktivieren (Austritt)

**Szenario:** Ein Mitarbeiter verlässt das Unternehmen.

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Mitarbeiter**

**Schritt 2:** Suchen Sie den Mitarbeiter in der Liste

**Schritt 3:** Ändern Sie in Spalte **E (Aktiv)** den Wert von `Ja` auf `Nein`

```
Vorher:  | 1002 | Schmidt | Anna | 01.07.2019 | Ja   | Innendienst |
Nachher: | 1002 | Schmidt | Anna | 01.07.2019 | Nein | Innendienst |
```

**Schritt 4:** Speichern Sie die Datei

#### Was passiert?

- ✅ Mitarbeiter erscheint nicht mehr in Dropdown-Listen für neue Ausgaben
- ✅ Historische Ausgabedaten bleiben erhalten
- ✅ Mitarbeiter erscheint nicht mehr in der Übersicht
- ✅ Restanspruch kann weiterhin abgefragt werden (für Dokumentation)

#### Wichtig

**Mitarbeiter NICHT löschen!** Durch das Löschen gehen alle historischen Daten verloren. Verwenden Sie immer das Aktiv-Flag.

---

### Fall 6: Neuen Artikel zum Sortiment hinzufügen

**Szenario:** Ein neues Bekleidungsstück wird in das Sortiment aufgenommen.

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Sortiment**

**Schritt 2:** Klicken Sie in die erste leere Zeile der Tabelle

**Schritt 3:** Geben Sie die Artikeldaten ein:

| Spalte | Feld | Beispiel 1 (jährlich) | Beispiel 2 (rollierend) |
|--------|------|----------------------|-------------------------|
| A | ArtikelID | 6 | 7 |
| B | Artikelname | T-Shirt | Winterjacke |
| C | AnspruchMenge | 3 | 1 |
| D | ZyklusJahre | 1 | 5 |
| E | ZyklusTyp | Kalender | Rollierend |
| F | Aktiv | Ja | Ja |
| G | Groessen | S,M,L,XL | S,M,L,XL,XXL |

**Schritt 4:** Speichern Sie die Datei

**Schritt 5:** Führen Sie `BtnUebersichtAktualisieren_Click` aus, um die Übersicht zu aktualisieren

#### Zyklus-Typen erklärt

| Typ | Bedeutung | Beispiel |
|-----|-----------|----------|
| **Kalender** | Anspruch gilt pro Kalenderjahr | 4 Hemden pro Jahr |
| **Rollierend** | Anspruch gilt X Jahre ab letzter Ausgabe | 1 Hoodie alle 3 Jahre |

---

### Fall 7: Anspruchsmengen anpassen

**Szenario:** Die Anspruchsmenge für einen Artikel soll geändert werden.

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Sortiment**

**Schritt 2:** Suchen Sie den Artikel in der Liste

**Schritt 3:** Ändern Sie den Wert in Spalte **C (AnspruchMenge)**

```
Vorher:  | 3 | Polo Shirt | 2 | 1 | Kalender | Ja |
Nachher: | 3 | Polo Shirt | 3 | 1 | Kalender | Ja |
```

**Schritt 4:** Speichern Sie die Datei

#### Auswirkungen

- ✅ Neue Anspruchsmenge gilt sofort für alle Berechnungen
- ✅ Bereits getätigte Ausgaben bleiben unverändert
- ⚠️ Restanspruch wird automatisch neu berechnet

---

### Fall 8: Ausgaben nach Datum sortieren

**Szenario:** Die Ausgabeliste soll chronologisch sortiert werden.

#### Schritt-für-Schritt

**Schritt 1:** Drücken Sie **Alt + F8**

**Schritt 2:** Wählen Sie `BtnAusgabenSortieren_Click`

**Schritt 3:** Klicken Sie **Ausführen**

**Ergebnis:** Die Ausgaben werden nach Datum sortiert (neueste zuerst).

#### Alternative: Manuell sortieren

1. Wechseln Sie zum Blatt **Ausgaben**
2. Klicken Sie auf den Dropdown-Pfeil in der Spalte **Datum**
3. Wählen Sie **Nach Datum sortieren (absteigend)**

---

### Fall 9: Innendienst-Mitarbeiter korrekt anlegen

**Szenario:** Ein neuer Innendienst-Mitarbeiter wird angelegt. Er soll nur 2 Hemden/Blusen erhalten (statt 4).

#### Schritt-für-Schritt

**Schritt 1:** Wechseln Sie zum Blatt **Mitarbeiter**

**Schritt 2:** Legen Sie den Mitarbeiter an (siehe [Fall 4](#fall-4-neuen-mitarbeiter-anlegen))

**Schritt 3:** Wählen Sie in Spalte **F (Bereich)** den Wert `Innendienst`

```
| 1005 | Becker | Julia | 01.03.2025 | Ja | Innendienst | Buchhaltung |
```

**Schritt 4:** Speichern Sie die Datei

#### Überprüfung

**Schritt 5:** Prüfen Sie den Restanspruch für diesen Mitarbeiter

**Erwartetes Ergebnis:**

| Artikel | Standard | Effektiv | Status |
|---------|----------|----------|--------|
| Hemd | 4 | **2** | Verfügbar |
| Bluse | 4 | **2** | Verfügbar |
| Polo Shirt | 2 | 2 | Verfügbar |

Die Spalte **Effektiv** zeigt 2 statt 4 für Hemd und Bluse.

#### Hintergrund

Diese Sonderregel ist in der Konfiguration hinterlegt:
- Blatt **Config** → Parameter `InnendienstHemdAnspruch` = 2
- Kann bei Bedarf angepasst werden

---

### Fall 10: 3-Jahres-Artikel verstehen

**Szenario:** Verständnis der rollierenden Zyklen für Hoodie und Softshelljacke.

#### Wie funktioniert der rollierende Zyklus?

**Beispiel: Mitarbeiter erhält Hoodie am 15.06.2025**

| Jahr | Anspruch | Begründung |
|------|----------|------------|
| 2025 | 0 | Gerade erhalten |
| 2026 | 0 | Erst 1 Jahr vergangen |
| 2027 | 0 | Erst 2 Jahre vergangen |
| 2028 | 1 | ✅ 3 Jahre vergangen, neuer Anspruch |
| 2029 | 1 | Anspruch noch nicht genutzt |
| 2030 | 1 | Anspruch noch nicht genutzt |

#### Wichtige Unterschiede zum Kalender-Zyklus

| Aspekt | Kalender (Hemd) | Rollierend (Hoodie) |
|--------|-----------------|---------------------|
| Anspruch verfällt | Am 31.12. des Jahres | Nie (bis zur nächsten Ausgabe) |
| Berechnung | Pro Kalenderjahr | Ab letzter Ausgabe |
| Typische Artikel | Hemden, Polo | Hoodie, Jacken |

#### Restanspruch-Anzeige interpretieren

**Noch nie ausgegeben:**
```
| Hoodie | 1 | 1 | 0 | 1 | Verfügbar (noch nie ausgegeben) |
```

**Zyklus noch nicht abgelaufen:**
```
| Hoodie | 1 | 1 | 0 | 0 | Nächste: 2028 |
```

**Neuer Anspruch verfügbar:**
```
| Hoodie | 1 | 1 | 0 | 1 | Verfügbar (letzte: 2025) |
```

---

## 3. Tipps und Best Practices

### Tägliche Arbeit

✅ **Ausgaben zeitnah erfassen**
- Erfassen Sie Ausgaben möglichst am Tag der Übergabe
- Vermeidet Fehler durch vergessene Einträge

✅ **Vor Ausgabe Restanspruch prüfen**
- Prüfen Sie den Restanspruch bevor Sie Kleidung ausgeben
- Verhindert Überschreitung der Kontingente

✅ **Regelmäßig speichern**
- Speichern Sie nach jeder Eingabe
- Nutzen Sie auch **Strg + S**

### Datenqualität

✅ **Einheitliche Schreibweise**
- Größen: Immer S, M, L, XL, XXL (keine Varianten wie "small")
- Bereich: Immer "Innendienst" oder "Außendienst"

✅ **Keine Zeilen löschen**
- Mitarbeiter deaktivieren statt löschen
- Ausgaben nicht löschen (ggf. Storno-Eintrag mit negativer Menge)

### Jahreswechsel

✅ **Übersicht aktualisieren**
- Nach dem Jahreswechsel die Übersicht für das neue Jahr aktualisieren
- Altes Jahr archivieren (Kopie der Datei)

---

## 4. Häufige Fragen (FAQ)

### Kann ich eine falsche Ausgabe korrigieren?

**Ja.** Sie haben zwei Möglichkeiten:

1. **Korrektur:** Ändern Sie die Werte direkt im Blatt **Ausgaben**
2. **Storno:** Erfassen Sie einen neuen Eintrag mit negativer Menge

### Warum zeigt der Innendienst-Mitarbeiter 4 Hemden statt 2?

Prüfen Sie:
1. Ist der Bereich korrekt auf "Innendienst" gesetzt?
2. Ist der Parameter `InnendienstHemdAnspruch` im Config-Blatt vorhanden?
3. Wurde die Restanspruch-Berechnung ausgeführt?

### Kann ein Mitarbeiter mehr erhalten als sein Anspruch?

**Ja**, das System warnt Sie, erlaubt aber die Eingabe. Die Warnung erscheint im Eingabedialog.

### Wie exportiere ich die Daten?

Die Daten können über Excel exportiert werden:
1. Blatt auswählen
2. **Datei** → **Speichern unter**
3. Format wählen (z.B. CSV, PDF)

### Was passiert bei einem Excel-Absturz?

Excel erstellt automatisch Wiederherstellungsdateien. Beim nächsten Start werden Sie gefragt, ob Sie diese wiederherstellen möchten.

### Können mehrere Personen gleichzeitig arbeiten?

**Nicht empfohlen.** Excel-Dateien sind für Einzelnutzung konzipiert. Bei gleichzeitiger Nutzung:
- Speichern Sie auf einem Netzlaufwerk
- Koordinieren Sie die Zugriffe
- Alternativ: SharePoint/OneDrive mit Co-Authoring

---

## Anhang: Kurzreferenz

### Tastenkürzel

| Kürzel | Aktion |
|--------|--------|
| Alt + F8 | Makro-Liste öffnen |
| Alt + F11 | VBA-Editor öffnen |
| Strg + S | Speichern |
| Strg + Z | Rückgängig |
| Strg + F | Suchen |

### Artikelliste (Standard)

| ID | Artikel | Anspruch | Zyklus |
|----|---------|----------|--------|
| 1 | Hemd | 4 (2 für Innendienst) | Jährlich |
| 2 | Bluse | 4 (2 für Innendienst) | Jährlich |
| 3 | Polo Shirt | 2 | Jährlich |
| 4 | Hoodie | 1 | 3 Jahre |
| 5 | Softshelljacke | 1 | 3 Jahre |

---

*Letzte Aktualisierung: Januar 2025*
