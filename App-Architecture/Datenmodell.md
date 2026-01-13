# Datenmodell - Bekleidungsverwaltung

## Übersicht der Tabellenblätter

```
┌─────────────────┐     ┌─────────────────┐
│   Mitarbeiter   │     │    Sortiment    │
│  (Stammdaten)   │     │   (Artikel)     │
└────────┬────────┘     └────────┬────────┘
         │                       │
         │    ┌──────────────┐   │
         └───►│   Ausgaben   │◄──┘
              │ (Erfassung)  │
              └──────┬───────┘
                     │
         ┌───────────┴───────────┐
         ▼                       ▼
┌─────────────────┐     ┌─────────────────┐
│   Uebersicht    │     │  Restanspruch   │
│  (Auswertung)   │     │   (Abfrage)     │
└─────────────────┘     └─────────────────┘
```

---

## Tabellenblatt: Mitarbeiter

### Struktur

| Spalte | Name | Datentyp | Pflicht | Beschreibung |
|--------|------|----------|---------|--------------|
| A | Personalnummer | Long | Ja | Eindeutige ID |
| B | Nachname | String | Ja | Nachname |
| C | Vorname | String | Ja | Vorname |
| D | Eintrittsdatum | Date | Ja | Datum des Eintritts |
| E | Aktiv | Boolean | Ja | Ja/Nein |
| F | Abteilung | String | Nein | Optional für Filter |

### Beispieldaten

| Personalnummer | Nachname | Vorname | Eintrittsdatum | Aktiv | Abteilung |
|----------------|----------|---------|----------------|-------|-----------|
| 1001 | Müller | Hans | 01.01.2020 | Ja | Vertrieb |
| 1002 | Schmidt | Anna | 15.03.2021 | Ja | IT |
| 1003 | Weber | Peter | 01.06.2019 | Nein | Produktion |

### Benannte Bereiche

- `tbl_Mitarbeiter` - Gesamte Tabelle (als ListObject)
- `lst_MitarbeiterAktiv` - Aktive Mitarbeiter für Dropdown

---

## Tabellenblatt: Sortiment

### Struktur

| Spalte | Name | Datentyp | Pflicht | Beschreibung |
|--------|------|----------|---------|--------------|
| A | ArtikelID | Integer | Ja | Eindeutige Artikel-ID |
| B | Artikelname | String | Ja | Bezeichnung |
| C | AnspruchMenge | Integer | Ja | Kontingent pro Zyklus |
| D | ZyklusJahre | Integer | Ja | 1 = jährlich, 3 = alle 3 Jahre |
| E | ZyklusTyp | String | Ja | "Kalender" oder "Rollierend" |
| F | Aktiv | Boolean | Ja | Artikel verfügbar? |

### Standarddaten

| ArtikelID | Artikelname | AnspruchMenge | ZyklusJahre | ZyklusTyp | Aktiv |
|-----------|-------------|---------------|-------------|-----------|-------|
| 1 | Hemd | 4 | 1 | Kalender | Ja |
| 2 | Polo Shirt | 2 | 1 | Kalender | Ja |
| 3 | Hoodie | 1 | 3 | Rollierend | Ja |
| 4 | Softshelljacke | 1 | 3 | Rollierend | Ja |

### Benannte Bereiche

- `tbl_Sortiment` - Gesamte Tabelle
- `lst_Artikel` - Artikelnamen für Dropdown

---

## Tabellenblatt: Ausgaben

### Struktur

| Spalte | Name | Datentyp | Pflicht | Beschreibung |
|--------|------|----------|---------|--------------|
| A | AusgabeID | Long | Auto | Fortlaufende Nummer |
| B | Datum | Date | Ja | Ausgabedatum |
| C | Personalnummer | Long | Ja | FK zu Mitarbeiter |
| D | MitarbeiterName | String | Formel | SVERWEIS auf Mitarbeiter |
| E | ArtikelID | Integer | Ja | FK zu Sortiment |
| F | Artikelname | String | Formel | SVERWEIS auf Sortiment |
| G | Menge | Integer | Ja | Ausgegebene Stückzahl |
| H | Kalenderjahr | Integer | Formel | =JAHR(B:B) |
| I | Bemerkung | String | Nein | Optional |

### Validierung

- **Personalnummer**: Dropdown aus `lst_MitarbeiterAktiv`
- **ArtikelID**: Dropdown aus `lst_Artikel`
- **Menge**: Ganzzahl > 0, max. 10
- **Datum**: Nicht in der Zukunft

### Benannte Bereiche

- `tbl_Ausgaben` - Gesamte Erfassungstabelle

---

## Tabellenblatt: Uebersicht

### Struktur (Matrix)

Dynamische Pivot-ähnliche Ansicht:

| | 2025 Hemd | 2025 Polo | 2025 Hoodie | 2025 Softshell | 2024 Hemd | ... |
|---|---|---|---|---|---|---|
| Müller, Hans | 2 | 1 | 0 | 1 | 4 | ... |
| Schmidt, Anna | 3 | 2 | 1 | 0 | 2 | ... |

### Formeln

Aggregation mit SUMMEWENNS:

```
=SUMMEWENNS(
    tbl_Ausgaben[Menge],
    tbl_Ausgaben[Personalnummer], $A2,
    tbl_Ausgaben[ArtikelID], B$1,
    tbl_Ausgaben[Kalenderjahr], $B$1
)
```

---

## Tabellenblatt: Restanspruch

### Eingabebereich

| Zelle | Beschreibung |
|-------|--------------|
| B2 | Abfragejahr (Dropdown 2025-2030) |
| B3 | Mitarbeiter (Dropdown aus lst_MitarbeiterAktiv) |
| B5 | Button "Berechnen" |

### Ausgabebereich

| Artikel | Anspruch | Ausgegeben | Rest | Status |
|---------|----------|------------|------|--------|
| Hemd | 4 | 2 | 2 | ✓ |
| Polo Shirt | 2 | 2 | 0 | Erschöpft |
| Hoodie | 1 | 0 | 1 | ✓ (zuletzt: 2022) |
| Softshelljacke | 1 | 1 | 0 | Erschöpft (nächst: 2028) |

### Berechnungslogik

#### Jährliche Artikel (ZyklusJahre = 1)

```
Restanspruch = AnspruchMenge - SUMMEWENNS(Ausgaben im Kalenderjahr)
```

#### Rollierende Artikel (ZyklusJahre = 3)

```
1. Finde letztes Ausgabedatum für diesen Artikel + Mitarbeiter
2. Wenn (Abfragejahr - Jahr(LetzteAusgabe)) >= 3:
      Restanspruch = AnspruchMenge
   Sonst:
      Restanspruch = 0
      NächsterAnspruch = Jahr(LetzteAusgabe) + 3
3. Wenn keine Ausgabe vorhanden:
      Restanspruch = AnspruchMenge (erster Anspruch)
```

---

## Tabellenblatt: Config (optional)

### Anwendungseinstellungen

| Parameter | Wert | Beschreibung |
|-----------|------|--------------|
| StartJahr | 2025 | Erstes Jahr der Erfassung |
| MaxZeilenAusgaben | 10000 | Performance-Limit |
| AutoAktualisieren | Ja | Übersicht automatisch updaten |

---

## Beziehungen

```
Mitarbeiter (1) ──────< (n) Ausgaben
     │
     └── Personalnummer (PK) = Ausgaben.Personalnummer (FK)

Sortiment (1) ──────< (n) Ausgaben
     │
     └── ArtikelID (PK) = Ausgaben.ArtikelID (FK)
```

---

## Datenintegrität

### Constraints (via Datenvalidierung)

1. **Personalnummer in Ausgaben** muss in Mitarbeiter existieren
2. **ArtikelID in Ausgaben** muss in Sortiment existieren
3. **Menge** muss positiv sein
4. **Datum** darf nicht in der Zukunft liegen

### VBA-Validierung

```vba
Public Function ValidateAusgabe(dtDatum As Date, _
                                lngPersonalnummer As Long, _
                                intArtikelID As Integer, _
                                intMenge As Integer) As Boolean
    ' Prüfungen...
End Function
```
