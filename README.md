# CicloRank

**Free, offline, open-source Windows desktop application for race management, manual lap timing, and instant Excel/PDF standings.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11%20(64--bit)-lightgrey.svg)]()
[![Support on Ko-fi](https://img.shields.io/badge/Ko--fi-Support%20this%20project-FF5E5B.svg)](https://ko-fi.com/alessio75vc)

CicloRank is built for cycling clubs, race officials, and event organizers-especially small associations and volunteer teams without access to expensive electronic transponder systems. It lets you import rider start lists, record finish-line crossings manually via keyboard, fix typos on the fly without stopping the clock, and generate official Excel and PDF results in seconds.

---

## Download (Ready-to-Use for Windows)

If you want to run CicloRank immediately without compiling the source code:

1. Open the **[Latest Release page](https://github.com/alessio75/CicloRank/releases/latest)**.
2. Download the `.zip` archive from the **Assets** section.
3. Extract all files into a folder on your PC and double-click `CicloRank.exe`.

### Note on Windows SmartScreen
Because CicloRank is an independent open-source project without a paid commercial code-signing certificate, Windows 10/11 may show a *"Windows protected your PC"* SmartScreen warning on first launch. Click **More info** and then **Run anyway** to start the program.

---

## Key Features

- **100% Offline & Crash-Safe:** Runs completely offline with zero telemetry. User preferences and an automatic crash-recovery backup (`sessione_temp.json`) are saved locally in `%LocalAppData%\CicloRank`.
- **Keyboard Lap Timing & Live Preview:** Type a bib number and press `Enter` to log each lap crossing, with instant rider name and category lookup as you type. Unknown bib numbers not present in the start list are still recorded and flagged automatically.
- **On-the-Fly Error Correction:** Fix mistyped bib numbers during hectic bunch finishes directly from the top 3 recent-pass summary boxes or by editing cells inside the lap-by-lap timing grid—standings recalculate immediately.
- **Zero-Dependency Excel & CSV Import:** Loads `.xlsx`, `.csv`, and `.txt` start lists natively (no Microsoft Excel required) with automatic header detection across 11 languages and multi-sheet club ID resolution. Automatically assigns sequential bib numbers if the bib column is missing, and skips/reports any duplicate bib numbers found in the file.
- **Instant Excel (`.xlsx`) & A4 PDF Export:** Generates styled Excel workbooks with individual lap splits and print-ready A4 PDF classification sheets grouped by category—or automatically rendered as a single overall standings table if no riders have a category assigned (via headless Microsoft Edge, Chrome, or Brave, with `.html` fallback).
- **Session Export/Import, Quick Heat Reset & Themes:** Save or reload complete race snapshots in `.json` format, switch between Light Mode and Dark Mode at any time, or click **RESET DATA** to clear recorded passes, stopwatch time, and unknown riders while keeping the loaded start list in memory to start a new heat immediately without re-importing.
- **11 Languages Supported:** Full UI, report headers, and spreadsheet column recognition in **Italian, English, Spanish, French, German, Portuguese, Dutch, Polish, Simplified Chinese, Russian, and Japanese**.

---

## Quick Start (5 Steps)

1. **Set Race Info (Optional):** Click **RACE INFO** to enter Federation/Committee, Organizer, Event Name, Location, Date, and Sport.
2. **Load Start List:** Click **LOAD RIDERS** to import an `.xlsx`, `.csv`, or `.txt` participant file (missing bib columns are auto-numbered sequentially and duplicate bibs are skipped and reported), or start timing directly without a pre-loaded list.
3. **Start & Record Laps:** Click **Start** on the stopwatch, type each rider's bib number in **ENTER BIB NUMBER**, and press `Enter`.
4. **Verify & Correct:** Edit any mistyped bib number in the top 3 quick-edit boxes or directly in the left lap table; filter live rankings by category on the right.
5. **Save, Export & Next Heat:** Save the race session (`.json`), click **EXPORT EXCEL** / **EXPORT PDF** to generate the official standings, or use **RESET DATA** to clear timing data and start a new heat with the same loaded riders.

---

## Frequently Asked Questions (FAQ)

**Does CicloRank require an internet connection at the finish line?**
No. Everything runs locally on your Windows 10/11 (64-bit) PC, including `.xlsx` generation and PDF rendering.

**Do I need Microsoft Excel installed?**
No. CicloRank reads and writes `.xlsx` OpenXML files and `.csv`/`.txt` lists natively without external office suites.

**What happens if the app or PC closes accidentally during a race?**
CicloRank automatically saves a temporary session file (`sessione_temp.json`) after every recorded or edited pass and prompts you to restore all race data on the next startup.

---

## Panoramica in Italiano

**CicloRank** è un software gratuito e open-source per Windows 10/11 (64-bit) pensato per piccole società sportive, giudici di gara e organizzatori che devono gestire il cronometraggio manuale al traguardo senza chip elettronici e senza connessione internet.

- **Download pronto all'uso:** Vai alla pagina **[Releases](https://github.com/alessio75/CicloRank/releases/latest)**, scarica il file `.zip`, estrailo in una cartella e avvia `CicloRank.exe` (al primo avvio, se compare l'avviso di Windows SmartScreen, clicca su *Ulteriori informazioni* → *Esegui comunque*).
- **Funzionalità principali:** Importazione iscritti da `.xlsx`, `.csv` e `.txt` senza bisogno di Excel installato (con numerazione progressiva automatica se manca la colonna pettorali e segnalazione di eventuali doppioni), rilevamento passaggi da tastiera con anteprima immediata dell'atleta, correzione rapida degli ultimi 3 passaggi o direttamente in tabella, azzeramento rapido (**AZZERA DATI**) per avviare nuove batterie mantenendo l'elenco iscritti in memoria, salvataggio automatico anti-crash, classifiche assolute e per categoria esportabili in `.xlsx` (con tutti i parziali giro per giro) e `.pdf` (con classifica assoluta automatica se nessuna categoria è valorizzata), tema Chiaro/Scuro e interfaccia in 11 lingue.

---

## Support & License

CicloRank is open-source software released under the **MIT License** — see the [LICENSE](LICENSE) file for details.

If CicloRank helped you organize a race or saved you hours of spreadsheet work, consider supporting its development on Ko-fi:

**[Support CicloRank on ko-fi.com/alessio75vc](https://ko-fi.com/alessio75vc)**