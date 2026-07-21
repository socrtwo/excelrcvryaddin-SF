# Excel Recovery Add-in — cross-platform edition

Recover data from corrupt Microsoft Excel workbooks (`.xlsx` / `.xls`).

[![Live app](https://img.shields.io/badge/live-app-34d399?style=for-the-badge)](https://socrtwo.github.io/excelrcvryaddin-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/excelrcvryaddin-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/excelrcvryaddin-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/excelrcvryaddin-SF?style=for-the-badge&color=22d3ee)](LICENSE)

**Live app: <https://socrtwo.github.io/excelrcvryaddin-SF/>** — no install, no
upload. Your workbook is processed entirely in your browser and never leaves
your device. It also works offline and can be installed as a desktop/mobile
app (PWA).

Two implementations live in this repository:

1. **The cross-platform PWA** (`web/`) — remakes of the *portable* recovery
   strategies of the original add-in, running client-side on Windows, macOS,
   Linux, ChromeOS, Android and iOS.
2. **The classic Windows COM add-in** (VB.NET, `ExcelRecoveryAddin.sln`) — the
   original ribbon add-in, which drives Microsoft Excel itself for the
   Excel-powered strategies that cannot run in a browser.

## How the PWA works

Every file first goes through the **S2 File Identifier**: it reads magic
numbers rather than trusting the extension, separates concatenated or
embedded foreign files (each downloadable on its own), recommends the right
S2 recovery tool for mismatched types, and hands the actual Excel bytes to the
recovery strategies:

| Strategy | Honors | What it does |
| --- | --- | --- |
| **ZIP Repair — Try** | `Impl_ZipRepairTry` (Info-ZIP `zip -FF`) | Scans the byte stream for ZIP local file headers, re-inflates every salvageable entry with the fault-tolerant [Immortal Inflater](web/immortal-inflate.js) — even with a truncated or destroyed central directory — and repackages a fresh, well-formed `.xlsx`. |
| **Extract Data (non-MS)** | `Impl_NonMSXlsxExtractData` / `…2` (doctotext, coffec) | Parses `xl/sharedStrings.xml` and `xl/worksheets/*.xml` directly — no Excel needed — and writes one CSV per sheet, with sheet names resolved from the workbook relationships. |
| **Salvage Cells** | ExcelFix-style extraction | Regex-harvests cell references, values and text runs from partially readable XML when the sheets no longer parse. |
| **.xls Text Rescue** | legacy BIFF best-effort | Printable-string sweep (ASCII + UTF-16LE) of an OLE2/BIFF workbook that nothing else will open. |

Outputs: a repaired `.xlsx` where the container can be rebuilt, plus `.csv`
per sheet and `.txt` dumps for the salvage paths.

### Windows add-in (Excel-powered strategies)

The remaining strategies of the original add-in automate Excel over COM and
therefore need **Windows with Excel 2003 or later** — they are intentionally
*not* part of the PWA:

- **Open & Repair** / **Open & Extract Data** (`xlRepairFile` / `xlExtractData` corrupt-load modes)
- **Open in Safe Mode** (`excel.exe /s /r`) and **Manual Calculations**
- **External References** — recover values via link formulas into a fresh workbook
- **Save As SYLK** / **Save As HTML** round-trip filtering
- **Open with WordPad** / **Excel Viewer** last-ditch viewing

Build it from source: open `ExcelRecoveryAddin.sln` in Visual Studio
(2010 or later), pick the Excel 2003 or 2007+ project, and build. The add-in
targets .NET Framework and the Office primary interop assemblies, so it must
be built on a machine with Office installed — which is why CI treats this
build as best-effort.

## Running the PWA locally

```sh
cd web
python3 -m http.server 8080
# open http://localhost:8080
```

Or simply open `web/index.html` in a browser — the app has zero external
dependencies (no CDNs, no network calls with your data).

## Development

- Engine + UI: `web/recovery.js` (the engine half is UMD — it also loads in
  Node for tests). Fault-tolerant unzip/DEFLATE: `web/immortal-inflate.js`.
- Tests: `node scripts/test-recovery.mjs` builds a real `.xlsx` fixture with
  Python's `zipfile`, corrupts it (truncated tail, zeroed central directory),
  and asserts the strategies still recover every cell.
- Syntax check: `node --check web/*.js`.
- If you change any cached asset, bump `CACHE` in `web/sw.js` or returning
  users will be served stale code.

## Releases

Platform bundles (Windows / macOS / Linux / ChromeOS / Android / iOS / Web)
are built by `scripts/build-releases.sh` and published automatically:

- Push a `v*` tag (`git tag v1.0.0 && git push origin v1.0.0`), **or**
- Actions → *Build & publish multi-platform releases* → *Run workflow*.

Each bundle ships the same static app plus a tiny platform launcher; verify
downloads with the attached `SHA256SUMS`. Build locally with
`bash scripts/build-releases.sh v1.0.0` (output in `dist/`).

## Repo map

```
web/                    the PWA (app, engine, service worker, manifest, icons)
Code/, Forms/, …        classic VB.NET COM add-in source
ExcelRecoveryAddin.sln  Visual Studio solution (Excel 2003 + 2007 projects)
scripts/                release packaging, icon generation, smoke tests
.github/workflows/      build.yml (CI), pages.yml (Pages deploy), release.yml
```

## SourceForge heritage

This project began as the **[Excel Recovery Add-In](https://sourceforge.net/projects/excelrcvryaddin/)**
on SourceForge — a VB.NET ribbon add-in coded by **Sergey Zhebka**, collecting
in one place Microsoft's documented Excel-corruption recovery methods plus
non-MS extraction methods powered by the doctotext, no-frills, coffec and
Info-ZIP projects. The original readme's advice still holds and shaped this
edition: *"It is probably wise to try to do a zip repair step on corrupt xlsx
files as a first step before trying to invoke other methods."*

The cross-platform edition remakes the portable strategies in JavaScript so
they can ship everywhere; the original add-in remains the Windows-native
offering. Migrated to GitHub via SF2GH Migrator.

## License

MIT — see [LICENSE](LICENSE).

*Maintained by [@socrtwo](https://github.com/socrtwo). Bug reports and pull
requests welcome at the [issue tracker](https://github.com/socrtwo/excelrcvryaddin-SF/issues).*
