# FormFiller

A local, offline smart form-filling tool: extract information from source files (résumés, info sheets, etc.) → intelligent matching → fill Word/Excel templates → produce files that keep the template's exact formatting.

**Key point**: pure rule-based implementation (alias tables + similarity matching). No LLM API calls, no network access, fully offline. All data stays in your local `data/` directory.

## Supported template types

| Template type | Example | Detection |
|------|------|------|
| `{{field}}` placeholder templates | employee onboarding form, quotation sheet | auto-detected placeholders |
| "label cell + blank cell" traditional forms | university application forms, job info sheets | auto-switches to form mode when 0 placeholders are parsed |

## Features

- **Data library**: upload documents (docx/docm/xlsx/pdf/txt/md/csv/json and images) to auto-extract key-value info; résumés whose text lives inside textboxes are also extracted; persisted in `data/db.json` — build once, reuse many times
- **Smart matching**: alias tables (`联系电话`→`电话`, `毕业院校`→`学校`, etc.) + suffix matching + character-overlap similarity; source and template fields don't need identical names
- **Derived fields**: age computed with month precision; composite labels (e.g. "最后学历毕业院校及学位") auto-joined from multiple source fields
- **Multi-row list blocks**: education / work experience / published papers / family members and other "header + data rows" tables, mapped by column-name similarity, with automatic row insertion
- **Citation parsing**: parse APA citations from résumé text (including textboxes), extracting author order, title, journal, ISSN, impact factor, indexing, year-volume(issue)-pages; built-in ISSN map for common journals; use `--set 本人=<English surname>` to auto-derive author order
- **Image placeholders**: `{{img:头像}}` auto-inserts images from the library; `头像_1.png`, `头像(2).png` numbered variants all match
- **Two ways to use**: CLI for one-shot results; Web UI (library + template management + generation history) for repeated use

## Quick start

### Install dependencies

```bash
pip install -r backend/requirements.txt
```

### CLI usage

```bash
# 1. Inspect a template (which placeholders / form fields it has)
python cli.py scan --template template.docx

# 2. Fill (multiple source files can be passed at once)
python cli.py fill --template template.docx --source resume.docx job_info.txt --out result.docx

# Fill missing fields with --set (repeatable)
python cli.py fill --template template.docx --source resume.docx --set 签名=John --out result.docx

# Form mode + paper table: specify your English surname to auto-derive author order
python cli.py fill --template application.docx --source resume.docx info.docm --set 本人=Wang --out application_filled.docx
```

When fields are missing, it **refuses to generate** by default and lists what's missing; add `--force` if leaving them blank is acceptable.

### Launch the Web UI

```bash
cd backend
python -m uvicorn main:app --host 127.0.0.1 --port 8765
```

Open http://127.0.0.1:8765 in your browser.

> Note: you must run from the `backend/` directory. The backend uses flat imports (`import extractor`), so starting from elsewhere raises `ModuleNotFoundError`.

## Run tests

```bash
# Module tests (50 cases: extraction, matching, filling, citation parsing)
cd backend
python test_modules.py

# API end-to-end tests (requires the server running)
python test_api.py
```

## Project structure

```
├── cli.py             CLI entry (scan / fill / serve)
├── backend/           FastAPI backend
│   ├── main.py        API service entry
│   ├── extractor.py   file info extraction (docx/docm/xlsx/pdf/txt/…, incl. textboxes)
│   ├── filler.py      placeholder template parsing & filling
│   ├── form_analyzer.py  table-form analysis, list blocks, citation parsing
│   ├── storage.py     local data & file storage
│   ├── make_samples.py   generate built-in samples
│   ├── test_modules.py   module-level tests
│   └── test_api.py       API end-to-end tests
├── frontend/          single-page web frontend
├── sample/            built-in samples (onboarding form, quotation sheet, fictional résumé, etc. — run the flow immediately)
└── data/              runtime data (library, templates, generated files; created on first run, not committed)
```

## Supported formats

- **Templates**: `.docx`, `.xlsx` (table-form mode currently supports `.docx` only)
- **Source files**: docx, docm (macro-enabled Word; content type normalized in memory on read), xlsx, pdf, txt, md, csv, json and common images
- **Legacy `.doc`**: the extractor doesn't read OLE2 binary directly. On Windows with Word installed, save as `.docx` first.

## Notes

- Generated Word/Excel files preserve the original template formatting as much as possible; complex charts, macros, and special controls may not be preserved (python-docx / openpyxl limitations)
- Company name and similar fields may live in **headers/footers**; scan headers when validating
- Always **review the generated file** manually, especially signature, date, and amount fields
- The `data/` directory holds your real data and is excluded via `.gitignore`; it is never committed

> 中文说明见 [README_zh.md](README_zh.md)。
