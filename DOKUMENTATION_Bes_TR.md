# Dokumentation: Bes TR.xlsm

**Datei:** `F:\CursorProject_Python\Bes TR.xlsm` (ca. 7,0 MB, Excel-Makroarbeitsmappe)  
**Autor / letzte Bearbeitung:** ubes / Ulrich Besenfelder  
**Erstellt:** 12.01.2019 · **Zuletzt geändert:** 06.09.2026 (laut Dateieigenschaften)  
**Stand dieser Analyse:** 13.09.2026

`Bes TR` ist die Steuerzentrale für das **Archiv Trampolin**: Personen, Wettkämpfe, Ergebnislisten (Designs), Fotos/Videos mit IDs, Windows-Verknüpfungen und Statistik. Die Arbeitsmappe ist kein reines Tabellenkalkulationsblatt, sondern eine VBA-Anwendung mit klickbaren Zellen als Oberfläche.

---

## 1. Zweck

Die Datei verbindet drei Welten:

1. **Excel-Daten** (Personen, Starts, Termine, Ergebnis-Layouts)
2. **Dateisystem** unter `F:\Archiv Trampolin …` (Events, Leute, ClubsNations, Register)
3. **Hilfsprogramme** (`es.exe`, `exiftool.exe`, VBScript unter `prog\helpers`)

Typische Aufgaben:

- Ergebnislisten als **Designs** auf Blatt T4 zeichnen, als JPG exportieren und in Event- und Personenordner kopieren
- **Personenstammdaten** (T5 / PersonData.txt) mit Ordnernamen und Link-Beschriftungen synchron halten
- **Starts** (CompetitorsList, T3) aus den Designs ableiten
- **Fotos/Videos** mit `p####-##` / `v####-##` versehen
- **.lnk**- und **VBS-Links** zwischen Personen, Vereinen, Nationen und Events pflegen und reparieren
- **Terminkalender** (T9) mit vorhandenen Event-Ordnern abgleichen
- Archiv-**Statistik** auf T1 aktualisieren

---

## 2. Dateisystem (erwartete Nachbarschaft)

Beim Öffnen setzt `Fill_ArrCPaths` die Pfade aus dem Speicherort der XLSM neu. Erwartete Struktur:

| ArrC | Pfad (aktueller Stand auf diesem Rechner) | Rolle |
|------|-------------------------------------------|--------|
| 1 | `F:\Archiv Trampolin prog` | Programm-Wurzel (übergeordnet zum Ordner `prog`) |
| 2 | `F:\Archiv Trampolin 1900-1999` | Archiv-Wurzel |
| 3 | `…\Events` | Wettkampf-/Event-Ordner (`YYYY-MM-DD Name Ort_Land`) |
| 4 | `…\Leute` | Personenordner |
| 5 | `…\Register` | Register-Links |
| 6 | `…\prog\helpers` | VBS, `FreePIDs.txt`, Tools |
| 7 | `…\Read me.txt` | Archiv-Dokumentation (Quellenkürzel u. a.) |
| 8 | `helpers\es.exe` | Everything-Suche |
| 9 | `helpers\exiftool.exe` | Metadaten in JPGs |
| 10 | `…\prog\Bes TR.xlsm` | Kopie-Pfad ohne Roaming |
| 11 | `…\ClubsNations` | Vereine / LTVs / Nationen |

Liegt die Datei unter einem Pfad mit `Roaming`, wird auf ArrC(10) umgebogen (Schutz vor OneDrive/Profil-Kopien). Beim Öffnen erscheint eine Warnung, wenn der Dateiname `xlsb` enthält.

---

## 3. Startverhalten

`Workbook_Open` (DieseArbeitsmappe):

1. **F3** auf `RunShowProcs` legen (Prozedurnamen suchen)
2. Ereignisse und Bildschirmaktualisierung kurz aus
3. Blatt **T1** aktivieren, Schalter „Code rechts“ togglen
4. `DoArrc` — Container **ArrC** aus T8 laden
5. `Fill_ArrCPaths` — Pfade 1–11 neu schreiben
6. Eintrag ins **Logbuch**

`Workbook_BeforeClose`: F4-Belegung zurücksetzen, Logbuch „closed“. (Kommentar und OnKey-Taste weichen voneinander ab: Start nutzt F3, Schließen setzt F4 zurück.)

Klick-Oberflächen arbeiten über `Worksheet_SelectionChange` (kein klassisches Ribbon-Menü). Viele Aktionen quittieren mit `Beep1`.

---

## 4. Blätter

Elf sichtbare Blätter. Namensschema: **T** = Arbeits-/Datenblätter, **W** = Werkzeuge.

| Blatt | Rolle | Datenumfang (Stand Analyse) |
|-------|--------|-----------------------------|
| **T1** | Startmenü, Archiv-Statistik, Datei-Kopie | Steuerzellen + Statistik |
| **T2** | Sammel-Updates, globale Namensänderung, Link-Reparatur | 6 Ankreuz-Boxen, „Do all“ / Stop |
| **T3** | CompetitorsList (Starts) | ca. **8 982** Datenzeilen, Kennzahl: 2 071 aktive Starts, 1 045 w / 1 026 m, 40 Nationen, 234 Clubs |
| **T4** | Designs / Ergebnislisten | ca. **256–263** Design-Titel in der Liste; große Zeichenfläche rechts |
| **T5** | PersonData | ca. **2 519** Namen (1 246 w / 1 273 m) |
| **T6** | CHECK / UPDATE (Wartungsmenü) | ~20 Aktions-Buttons in Spalte 2 |
| **T7** | Label-Photos (beschriftete Fotos) | Scrollbar, ToDo-Ordner, Foto-Vorschau |
| **T8** | Container **ArrC** | Indizes 1–120, Werte + Kommentar in Spalte E |
| **T9** | TerminKalender | ca. **626** Termine, davon **302** mit Event-Ordner |
| **W1** | CountDown / Link-Fortschritt | z. B. Good/Bad Links |
| **W4** | Entwickler-Inspektor | Module, References, Prozeduren, Shapes, Farben, SysInfo |

### 4.1 T1 — Startseite „Archiv Trampolin“

Klickbare Aktionen (über Textmarker-Zellen, nicht Formular-Buttons):

- Ordner öffnen: Events, ClubsNations, Leute, Register, ToDo
- Dateien anzeigen: Read me.txt, Logbuch.txt, Terminkalender.txt, Competitors.txt, PersonData.txt
- **Save a copy** der Arbeitsmappe (zuletzt: `Bes TR 20260815_090421.xlsm`)
- **Update Statistics** (füllt ArrC(32)/(33) und die Zähler rechts: Photos, Videos, Links, Unterordner je Bereich)

Anzeige-Schalter: Menü ein, reduzierte Höhe, Gitterlinien, Code rechts.

### 4.2 T2 — Sammeljobs

Ankreuzbare Blöcke:

1. Fortlaufende Nummern oberhalb der Designs (Dgs)
2. Namen einer Person **global** ändern (Vorname/Nachname alt → neu)
3. Alle Links reparieren (broken .lnk), mit Unterboxen Events / ClubsNations / Leute

Buttons: **Do all**, **Stop**. T2 zählt aktuell u. a. 317 Event-Ordner, 198 ClubsNations-Ordner, 2521 Leute-Ordner.

### 4.3 T3 — CompetitorsList

Spalten (ab Zeile 6): Nachname, Vorname, m/w, Alter, Club, LTV, Nation, Event-Titel aus T4-Design, Disziplin (Einzel / Synchron / Mannschaft), Rang.

Aktionen: Liste aktualisieren, Textdatei erzeugen, Suche (M11), Sortierung über Kopfzeile. Herkunft: Namen und Platzierungen aus den T4-Designs, Stammdaten-Anreicherung aus PersonData (`M_PD`).

### 4.4 T4 — Designs (Kern der Ergebnislisten)

Jedes Design ist ein gerahmter Bereich auf dem riesigen Blatt (Koordinaten in ArrC 36–40, 71–72, 90–102). Die **List of DesignTitles** (links) verweist auf:

- Titel, Zellbereich (z1/s1/z2/s2), Event-Ordnername, extrahierte Namen

Klicks und ActiveX:

- Sprung Titel → Design, „new“, LU/ub, Spaltenbreite SomeDgData
- **HelperBox** (ListBox Namensvorschläge) + **CellBox** (Texteingabe über der Zelle)
- CmdOpen, CmdCreate, CmdBack, Cmdinfo, CmdSingleRanking
- Namenssuche mit Filter; Geschlecht oft über Zellenfarbe (hellrot = w, hellblau = m)
- Export **Result-JPG** in den Event-Ordner; Kopien/Links in Personenordner

Layout-Familien (Beispiel): `E2NV1` = Einzel, 2 Namensspalten, Nation, Verein, Punkte.

### 4.5 T5 — PersonData

Kopf ab Zeile 6:

| Spalte | Inhalt |
|--------|--------|
| 3–4 | Nachname, Vorname |
| 5 | m/w (färbt die Zeile) |
| 6–9 | Jahrgang, Club, LTV, Nation |
| 10 | Personenordner-Name |
| 12–15 | Rufname (rn), Geburtsname (gn), weitere Nachnamen (vh), Schreibweise (sw) |
| 16–17 | Geburts-/Sterbedatum |
| 18 | Ordner öffnen (Wingdings) |
| 19 | Sprung nach T2 „Namen global ändern“ |
| 22 | Vorname Nachname |

Suche in J4/J5, Sortierung über Zeile 6. `M_Load` baut daraus Arrays für Link-Namen (`LiVnNn`, `LiNnVn`, `LiRn`) inkl. Club/Nation und „geb …“.

### 4.6 T6 — CHECK / UPDATE

Spalte-2-Menü (Auswahl):

- IDs in offenen Event-Ordnern setzen, LTV/Nation in T5 ergänzen
- Zeichenketten in Event-Ordnernamen ändern
- Pfade, PIDs/VIDs prüfen, private JPG-Tags entfernen
- fehlende Personen (Leute vs. T5)
- CompetitorsList, Dg-Positionen, Alter in Designs, Result.jpg aktualisieren
- Links Person–Club–LTV–Nation, Event-Links in Personenordner
- Dateinamen formatieren, DateSquares, Icons, Konsistenz, tote Links

Bestätigungsdialoge liegen weiter rechts (Ja/Nein/Abbrechen, Eventordner öffnen).

### 4.7 T7 — Label-Photos

Workflow: Verknüpfung eines Fotos in den **ToDo-Ordner** legen → Label-Foto erzeugen (Name, Event, Alter, pID/vID). Scrollbar blättert die Liste. Sprünge: Foto im Event-Ordner, im Leute-Ordner, Label-Archiv.

### 4.8 T8 — ArrC (globaler Zustand)

Ein eindimensionales Array `ArrC(1…n)`, gespiegelt in T8 Spalte D. Beim Öffnen: `Fill_ArrCFromT8`, danach Pfade 1–11 überschrieben. Weitere Indizes halten UI-Zustand (letzte Selektion T4, Sortierung T3/T5, UF2-Position, Countdown, Statistik-String, aktive Foto-ID auf T7, …). **Neue Werte:** Zeile mit fortlaufender Nummer in T8 ergänzen.

### 4.9 T9 — TerminKalender

Spalten: Datum (intern `YYYYMMDD` + Anzeige), Event, Ort, Region/LTV, Nation, Level, SubRegion, Bemerkung, Ordner existiert, Ordnername, Pfad.

Beim Aktivieren: neue Event-Ordner eintragen, Ordner-Icons, PepUp-Farben. Klick auf exist-Spalte öffnet den Ordner. „Knollen“ in Spalte A steuern Sammelaktionen. T9 speist `z Terminkalender.txt`.

### 4.10 W1 / W4

- **W1:** Fortschrittsanzeige bei langen Link-/Ordnerläufen. Enthält außerdem experimentelle Foto-Helfer (HEIC/DNG → JPG via ImageMagick) mit festen Pfaden unter `G:\Archiv Photos\…` — nicht Teil des Trampolin-Hauptpfads.
- **W4:** Entwicklerkasten: VBA-Module, Verweise, Prozeduren, Shapes, Farben, Prozesse. Nützlich zur Selbstdokumentation der Mappe.

---

## 5. UserForms

| Form | Aufgabe |
|------|---------|
| **UF1** | Dialog (OK), u. a. Person/Nation-Auswahl im T4-Kontext (`ShowUF1`) |
| **UF2** | Verschiebbares Overlay ohne Titelleiste (letzte Position ArrC 69/70) |
| **UF3** | Klammern/„Brac“-Namen (Zähler ArrC 87/88), Modul `M_Brac` / `M_UF3` |
| **UF4** | Modelless Fortschrittsbalken (`UpdateCountdown`) |
| **UF5** | Neues Event anlegen (Vorgabe `1999-12-31 XY-Wettkampf Ort/CH`, Zeilen/Spalten) |
| **UF6** | Wahl m/w |

---

## 6. VBA-Architektur

Ca. **74** Code-Komponenten, ca. **1 550** Subs/Functions (inkl. Tests `zzz_*` / `ooo_*` und WinAPI in M99).

Konventionen:

- `zzz_Modulname` — Sprungbrett / `showProcs`
- `Option Explicit` durchgängig
- `EE 0/1` — `EnableEvents` aus/an
- `DoArrc` / `FillArrC` — Zustand
- `LogBuch` — Protokoll
- Klicklogik in den **Sheet-Modulen**, Fachlogik in **M_*** und nummerierten **Mnn**

### 6.1 Fachmodule (Auswahl)

| Modul | Thema |
|-------|--------|
| M_Fill | ArrC füllen, Pfade, Sex-Farben, IgnoreThisName |
| M_Load | T5-Spalten und 17-Spalten-Personenarray |
| M_PD | PersonData.txt lesen/schreiben, Alter in Competitors |
| M_Update | Sammel-Update (Jpg, Links, Competitors) |
| M_Link / M_LinkVbs / M_Vbs | .lnk und VBS-Links erzeugen/reparieren |
| M_Check / M_Check2 | Konsistenz, Explorer-Ansicht |
| M_Brac | Klammer-Namen auf Designs |
| M_Icon | ICO aus JPG |
| M_Label / M_T7 | Label-Fotos |
| M_T4DgData / M_T6 / M_T8 / M_T9 | Blatt-spezifische Helfer |
| M_HelperBox | T4 Namens-ListBox |
| M05 | Terminkalender-Text |
| M06 | Personenordner, Result-JPG-Kopien |
| M11 | Statistik, Read-me-Quellen, Dateisuche (91 Prozeduren) |
| M98 / M99 | Datei/Ordner/API, Explorer, Encoding, Referenzen |
| M41–M48, M61–M63 | Designs, Ranglisten, Dateisystem-Batches |

Extrahierter VBA-Quelltext zur Analyse liegt unter `_bes_tr_analysis\` (nicht Teil der XLSM selbst).

---

## 7. Datenflüsse (vereinfacht)

```
Event-Ordner (Fotos, Videos, Result.jpg)
        ↑ JPG / Links
T4 Designs ──► CompetitorsList (T3) ──► Competitors.txt
        ↑ Namen
T5 PersonData ◄──► Leute-Ordner, PersonData.txt
        ↓
ClubsNations / Register  (VBS-Links)
        ↓
T9 Terminkalender ◄──► z Terminkalender.txt
        ↓
T1 Statistik / Read me.txt / Logbuch
```

IDs:

- **pID** `p####-##` — Foto; freie IDs in `helpers\FreePIDs.txt`
- **vID** `v####-##` — Video
- Dateinamen von Fotos erwarten oft ein **Quellenkürzel** (`… [aa].jpg`); ohne Kürzel keine automatische pID (Ausnahme `a Leute.jpg`)

---

## 8. Bedienung (Kurz)

| Taste / Ort | Wirkung |
|-------------|--------|
| F3 | `RunShowProcs` — Prozedurnamen suchen |
| Klick auf beschriftete Zellen | Menüaktionen (T1, T6, T2, …) |
| T4 List of DesignTitles | Sprung ins Layout; Kopf aktualisiert die Liste |
| T5 Spalte 18 / 19 | Ordner öffnen / Namensänderung T2 |
| T9 Spalte 13 | Event-Ordner öffnen, wenn vorhanden |

Lange Läufe: Stop auf T2, Countdown auf W1, Fortschritt UF4.

---

## 9. Abhängigkeiten und Risiken

- **Makros müssen aktiv** sein; ohne VBA ist die Mappe nur Daten.
- Feste Laufwerke **F:** (Archiv) und teils **G:** (private Foto-Helfer auf W1).
- Externe Tools: Everything (`es.exe`), ExifTool, optional ImageMagick, Windows Script Host.
- `CreateObject("Scripting.FileSystemObject")`, `WScript.Shell` (Shortcuts), `Shell.Application` (Explorer).
- M99: WinAPI (`ShellExecute`, Speicher, Prozessliste).
- Roaming/OneDrive-Kopien werden erkannt, können aber Pfade und Links zerlegen.
- `GetPID` in T6 schreibt `Get_NextFreePID` — inkonsistenter Funktionsname (Fehlerquelle beim Aufruf aus VBS).
- Named Range `Test1` zeigt `#REF!`.
- Personen- und Ergebnisdaten sind umfangreich; Kopien der XLSM per T1 „Save a copy“ anlegen, bevor Massen-Updates (T2/T6) laufen.

---

## 10. Kennzahlen (aus der geöffneten Mappe)

| Kennzahl | Wert |
|----------|------|
| VBA-Komponenten | 74 |
| Subs/Functions (gezählt) | ~1 550 |
| T3 Starts (Zeilen) | ~8 982 |
| T4 Design-Titel | ~256 (Liste) / ~263 Zeilen in Spalte 2 |
| T5 Personen | ~2 519 |
| T9 Termine | ~626 (302 mit Ordner) |
| Event-Ordner (T2) | 317 |
| ClubsNations-Ordner (T2) | 198 |
| Leute-Ordner (T2) | 2 521 |
| ActiveX-Steuerelemente in der Datei | 8 |
| UserForms | 6 |

---

## 11. Weiterarbeit

Sinnvolle nächste Schritte, falls die Mappe modernisiert oder nach Python/Streamlit (`app.py` im gleichen Ordner) überführt werden soll:

1. ArrC und Pfadlogik als Konfiguration (JSON) herausziehen
2. T5/T3/T9 als Tabellen (CSV/SQLite) exportieren
3. T4-Designs als eigenes Layout-Format beschreiben (Gruppen E1VN1, M2V8, …)
4. Link-Reparatur und ID-Vergabe als Skripte ohne Excel
5. Die vorhandenen `zzz_`-Einstiege und `showProcs` als lebendiges Inhaltsverzeichnis nutzen
