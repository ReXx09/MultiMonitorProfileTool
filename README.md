# Multi-Monitor Profile Tool

Ein PowerShell WPF-Tool zur Verwaltung von Multi-Monitor-Layouts unter Windows 11. Erstelle Profile für verschiedene Setups (z. B. Arbeiten, Streaming, Gaming) und weise jedem Programm eine feste Zone auf einem bestimmten Monitor zu.

---

## Features

- **Profilverwaltung** — Erstelle und verwalte beliebig viele Layout-Profile
- **Drag-and-Drop Layout-Editor** — Ziehe Programme per Maus auf Monitore und wähle eine Zone
- **Live-Zone-Preview** — Während des Ziehens wird die Zielzone live angezeigt (grüner Monitor-Highlight)
- **Automatische Wiederherstellung** — Beim Profilwechsel werden Fenster automatisch positioniert
- **Fenster schließen vor Profilwechsel** — Verhindert Überlagerungen durch automatisches Schließen alter Fenster
- **Windows Autostart** — Tool kann beim Windows-Start automatisch gestartet werden
- **Per-Monitor DPI Awareness V2** — Korrekte Skalierung bei unterschiedlichen DPI-Monitoren
- **DE / EN Lokalisierung** — Sprache in den Einstellungen umschaltbar
- **Debug-Logging** — Optionales Logging aller Aktionen in eine Logdatei

---

## Voraussetzungen

- Windows 10 / 11
- PowerShell 5.1 oder höher
- .NET Framework 4.x (in Windows integriert)
- Ausführungsrichtlinie: `RemoteSigned` oder `Bypass`

---

## Installation

1. Repository klonen oder ZIP herunterladen:
   ```powershell
   git clone https://github.com/ReXx09/MultiMonitorProfileTool.git
   ```

2. In den Ordner wechseln:
   ```powershell
   cd MultiMonitorProfileTool
   ```

3. Tool starten:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -STA -File .\MultiMonitorProfileTool.ps1
   ```

   Oder per Doppelklick auf `launcher.bat`.

---

## Erste Schritte

### Profile anlegen

1. Tab **Profile** öffnen → **Neues Profil** klicken
2. Profilnamen eingeben (z. B. `Coding`, `Gaming`, `Präsentation`)

### Layout zuweisen

1. Tab **Layout-Editor** öffnen
2. Das gewünschte Profil im Dropdown auswählen
3. Links werden alle offenen Programme angezeigt
4. Ein Programm per Maus auf den gewünschten Monitor ziehen
5. Zone auswählen (z. B. `LeftHalf`, `RightHalf`, `Fullscreen`, `TopLeft` ...)
6. Die Regel ist gespeichert und wird als Kachel im Canvas angezeigt

### Layout anwenden

- **Dashboard** → **Layout aktives Profil anwenden**
- Oder das Profil in der Profilübersicht auswählen → **Layout anwenden**

---

## Zone-Übersicht

| Zone | Beschreibung |
|------|-------------|
| `Fullscreen` | Gesamter Monitor |
| `LeftHalf` | Linke Hälfte |
| `RightHalf` | Rechte Hälfte |
| `TopHalf` | Obere Hälfte |
| `BottomHalf` | Untere Hälfte |
| `TopLeft` | Oberes linkes Viertel |
| `TopRight` | Oberes rechtes Viertel |
| `BottomLeft` | Unteres linkes Viertel |
| `BottomRight` | Unteres rechtes Viertel |

---

## Einstellungen

| Einstellung | Beschreibung |
|-------------|-------------|
| Layout nach Moduswechsel wiederherstellen | Wendet das Layout automatisch nach einem Anzeigemodus-Wechsel an |
| Fehlende Programme automatisch starten | Startet Programme, die im Profil definiert sind, aber nicht laufen |
| Verzoegerung (ms) | Wartezeit nach dem Moduswechsel vor dem Anwenden des Layouts |
| Startwartezeit (ms) | Wartezeit nach dem Starten eines Programms |
| Sprache | DE / EN |
| Mit Windows starten | Fügt das Tool zum Windows Autostart hinzu |
| Ausgeschlossene Prozesse | Systemprozesse, die nie erfasst oder verschoben werden |

---

## Autostart

Im Tab **Einstellungen** die Checkbox **Mit Windows starten** aktivieren und auf **Speichern** klicken. Das Tool trägt sich dann in den Registry-Key `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` ein und startet beim Login automatisch im Hintergrund über `launcher.bat`.

---

## Launcher

`launcher.bat` startet das Tool ohne sichtbares Konsolenfenster:

```batch
start /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -STA -File "%~dp0MultiMonitorProfileTool.ps1"
```

---

## Konfiguration

Die Konfiguration wird in `monitor-profiles.json` gespeichert (liegt im gleichen Verzeichnis wie das Skript). Die Datei enthält alle Profile, Fenster-Regeln und Einstellungen. Sie wird beim ersten Start automatisch erstellt.

---

## Lizenz

MIT License — frei verwendbar und anpassbar.
