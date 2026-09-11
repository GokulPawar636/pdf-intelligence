# PDF Intelligence

A Python CLI that extracts native PDF content and produces a validated Excel
workbook. Report selection, columns, rows and measurement counts follow the
input. The measurement summary follows the supplied sample's layout.

## Install and run

Python 3.11 is the tested runtime. Create a virtual environment and install the
pinned direct dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe pipeline_main.py PDF1.pdf --offline
```

## Streamlit UI

Choose **Excel** under Input format to import one `.xlsx` or `.xls` workbook.
The importer finds labelled Object/Feature, Control, Nominal/Tolerance and
Reading/Sample or dated measurement columns across sheets. It preserves source
order, missing readings, units and numeric precision, and generates the same
two-sheet report with live formulas. Limits are 100 MB, 500 MB expanded XLSX
content and 2 million cells. Recalculate and save source formulas in Excel
before uploading; the importer reads cached results and does not execute them.
Unknown schemas produce an explanation rather than guessed measurements.
Corrupt tolerance signs and missing specifications require review.

The UI has two input choices: **PDF** and **Excel**, followed by upload, preview
and download. Advanced controls are kept out of the upload flow. Programmatic
Excel imports still support validated manual mappings through `import_excel`.
Native feature-table PDFs may omit Pass/Fail cells; numeric Meas cells still
identify their rows. Numeric Excel inputs in scientific notation are supported.

Set `GROQ_ENABLED=true` and `GROQ_API_KEY` in `.env` to enable server-side AI
assistance for unfamiliar content. This sends extracted PDF content or Excel
sheet previews to Groq. AI is off by default. Optionally set
`GROQ_MODEL` (default `qwen/qwen3.8-27b`). Excel AI proposes only source column
indices, which are validated for bounds, overlap, and numeric readings. It never
generates replacement measurements. AI mappings are flagged for review. Known
Excel layouts use the local mapper without an API call. Network/provider errors
produce a retry message for Excel; PDF imports retain a marked local review
report when possible. Keep the key out of source control.

The UI exports a clearly marked review report by default when quality checks flag an issue.
Use CLI `--strict` to require strict validation.
Warnings remain visible in the UI and measurement workbook. This is not an accuracy
certification: no PDF converter can guarantee correct interpretation of every
scan, handwritten page, graphical symbol, or unfamiliar layout.

Page-level checks record native text, OCR-required pages, drawing-only pages,
and blank/sparse pages in each batch manifest. Mixed PDFs are checked page by
page so readable pages do not conceal scanned pages. OCR is not currently
installed or executed by this project; scanned content requires an OCR step.

Install the direct dependencies and start the upload interface:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

Upload one or multiple PDFs together. Conversion starts automatically, and a
formatted worksheet preview appears beside the upload panel. Select Summary
Report or Differences, then click **Download Excel** for the complete workbook.
Processing is local and requires no API key. Changing the uploads rebuilds the
report; changing the preview sheet does not repeat extraction.

Each PDF can have a different page count. Limits are 100 MB / 500 pages per PDF
and 500 MB / 2,500 pages per batch, with no separate fixed file-count limit.
The preview shows up to 200 rows and 80 visible columns; the download contains
the complete workbook. Online interpretation remains available through the CLI.

Comparable measurements are matched by feature, control, part, unit, reference,
alignment and specifications, rather than row position. Each PDF contributes a
measurement column (or more when it contains multiple readings). Missing
readings stay blank, and exact duplicate PDFs are skipped with a warning.
Insufficient identity keeps records separate. Unrecognized measurement layouts
produce a consolidated source-field report for review instead of guessed
statistics. Scanned PDFs without readable text require OCR before extraction.

Native feature tables with repeated Control / Nom / Meas / Tol / Dev / Test
headers also use a column-position reader when generic table detection fails.
It retains blank nominal cells, geometric-control readings, multi-number
balloons, and Part # / Serial # metadata. Graphical control symbols that are
absent from PDF text are labeled for review. Unsigned geometric tolerances do
not become invented specification limits. Angular readings without an explicit
unit remain separate across files. These cases retain their numeric readings
and formulas in the formatted measurement report.

Combined runs save `Combined_Report.xlsx`, a manifest, and per-PDF extraction
JSON under `data/output/combined/<run ID>/`. The workbook has two sheets:
Summary Report and Differences.

## Container deployment

The included container runs the upload UI as an unprivileged user and exposes a
health endpoint. Build and run it with Docker Compose:

```powershell
docker compose up --build
```

Open `http://localhost:8501`. Set `GROQ_API_KEY` through your deployment
platform's secret store when online interpretation is required; do not add it to
the image, source tree, or compose file. The CI workflow installs the pinned
dependencies, runs `pip check`, and executes the regression suite on every push
and pull request.

The CLI prints the exact workbook path. Each run receives its own directory:

```text
data/output/<input stem>/<UTC timestamp>_<run ID>/
  <input stem>_summary.xlsx  # or <input stem>.xlsx for a general report
  manifest.json
  intermediate/
    raw_document.json
    structured_document.json
    semantic_document.json
```

Repeated runs and same-named inputs are isolated. An open previous workbook does
not block a new run. Workbooks are built in a temporary file, checked for ZIP
integrity and cached Excel errors, then atomically published. A failed export
preserves the previous destination file and removes its temporary workbook.

Examples:

```powershell
# Automatically choose summary or general content sheets
python pipeline_main.py input.pdf --offline
# Require passing the implemented document/measurement quality checks
python pipeline_main.py input.pdf --offline --strict
# Explicit report selection
python pipeline_main.py input.pdf --offline --report summary
python pipeline_main.py input.pdf --offline --report dynamic
# Independent batch processing: one workbook per PDF
python pipeline_main.py first.pdf second.pdf --offline --output-dir reports
# Combine PDFs into one workbook, as in the upload UI
python pipeline_main.py first.pdf second.pdf --offline --combine
# Save page renders and embedded images for diagnosis
python pipeline_main.py input.pdf --offline --render-assets
```

The default limits are 100 MB and 500 pages, configurable with `--max-mb` and
`--max-pages`. Invalid, encrypted, empty and over-limit PDFs are rejected before
extraction. A batch continues after an input fails, and exits with code 1 if any
input failed. Combined mode validates all inputs first and fails without a
partial workbook if any input cannot be processed. Successful execution exits
with code 0.

## Measurement summary

Measurement workbooks contain exactly two sheets: Summary Report (the main
sheet) and Differences. The summary contains the title and measurement table,
without an overview block or raw data sheet. Only supplied readings get columns; no
empty historical dates or artificial measurements are added.

- Object, Control, Nominal, Lower/Upper Tolerance, Tolerance and measurement(s)
  are followed by Mean, Std Deviation, UCL, LCL, Min, Max and Remark.
- Sample standard deviation requires two or more readings. UCL/LCL are mean
  plus/minus three standard deviations. `N/A` is displayed when unavailable,
  with one explanatory note above the table. Zero variation from identical
  repeated readings is a valid numeric zero.
- Repeated readings are grouped only when explicit object, control, unit,
  specification, balloon, reference and available part/batch context agree.
  Records without sufficient identity are kept separate. Numbered reading,
  measurement and actual-value columns, and JSON measurement arrays, are supported.
- Signed symmetric/asymmetric tolerances produce limits. Explicit absolute
  limits stay absolute; conflicting specifications or unreadable measurements
  flag the row for review. Derived limit formulas reference numeric offsets in
  the audit columns instead of embedding different constants in each row.
- Remarks check the latest supplied reading against specification limits;
  they do not assert statistical process stability. Control limits and
  specification limits are distinct.
- Audit columns are collapsed to keep the main table readable. They contain
  source pages, balloons, references, source results, notes, units, evidence,
  reading counts and tolerance offsets. Full source records remain in the
  intermediate JSON files produced by the CLI.
- Formulas and cached values are both stored. Excel recalculates on opening.
  Edit readings, nominal values, and signed tolerances directly in the workbook.
  Insert entire reading columns between Tolerance and Mean (including immediately
  before Mean); statistics expand automatically. Keep the
  Tolerance and Mean anchor columns, and the formulas, intact.
- Supported editable tolerance text includes `+/-0.5`, `±0.5`, and `+0.5/-0.2`.
  Source absolute limits remain editable absolute values. Replacing a derived
  limit formula with a number deliberately overrides that limit.
- Differences contains one live table with one row per measured feature, using
  its latest reading, including future reading columns. The first two columns
  identify the measured feature and measurement type; source names are preserved.
  SD/UCL/LCL show
  `N/A` until at least two numeric readings exist; nonnumeric readings cause
  `Review`. Source review notes remain until verified and cleared in the
  expandable detail columns. The Mean header comment explains how to add columns.
- Unrelated columns outside the reading area do not count as measurements.
  New formulas for arbitrary business fields cannot be inferred safely. General
  document exports retain their source-defined columns and Excel tables.

## General documents

For independent CLI exports, `--report dynamic` produces an Overview, content-driven sheets and an Evidence
sheet. In auto mode, non-measurement tables use this format. Unrelated schemas
are kept separate. Headerless tables retain every row with positional columns.
Source text, repeated fields, identifiers and long values are preserved. Text
beginning with `=` is written as text, never executed as an Excel formula.

## Online interpretation

Offline mode sends no data to a model and retains tables, labeled values and
prose. For Groq interpretation, configure `GROQ_API_KEY` in `.env` (see
`.env.example`) and omit `--offline`. Choose a supported model with `--model`.
The SDK has a 30-second timeout; the extractor controls retries itself.

Model fields must match cited source text. Unsupported fields are discarded and
flagged, units require source evidence, and conservative local normalization
preserves identifiers. Original source text is kept alongside interpreted
records. Network or extraction failures fail the run; they are not reported as
successful empty workbooks. Live provider/model availability is not guaranteed
by the offline tests.

## Manifest and quality checks

`manifest.json` records the source SHA-256, run ID, stage, selected report, record
count, output path/hash and quality warnings. Status is `complete`,
`needs_review`, or `failed`. A completed run means processing succeeded, not
that every extracted relationship has been independently verified.

`--strict` prevents workbook publication for document warnings, empty extraction,
unsupported measurements or detected specification conflicts. It is not a proof
of extraction correctness. Source-level geometric/header heuristics still need
review for unfamiliar layouts. Formula caches are validated before publication;
opening and saving a workbook in Excel changes its file hash.

Record-level warnings are included in the final manifest and upload screen in
both offline and online modes. Inferred or uncertain table headers therefore
produce `needs_review`; strict mode blocks their publication until the source
layout is supported without those warnings. This also applies to the supplied
sample: its numeric regression checks pass, but its headers are inferred.

The manifest's `quality` section reports warning counts and measurement counts.
`accuracy_verified` remains false: software confidence and passing automated
checks are not a measured accuracy percentage. Establish deployment accuracy
using independently transcribed examples from each intended PDF layout,
including failed/scanned inputs, negative values, missing readings, and
specification conflicts. The regression suite checks those implemented cases;
it does not establish accuracy on unseen layouts.

AI evidence validation rejects numeric fragments and dropped signs. Invalid,
over-precision, or conflicting measurement specifications suppress tolerance
limits and require review instead of producing an apparently passing result.

## Verification and deployment limits

```powershell
python -m pytest -q
```

Regression tests cover the supplied PDF values, different measurement layouts,
invoices, inventory, prose, empty pages, statistical calculations, isolated
runs, locked output failures, atomic publication, encrypted/large inputs,
invalid measurements and evidence validation. `verify_summary.ps1` optionally
uses locally installed Windows Excel for recalculation and PDF previews; it is
not a runtime dependency and never enables macros.

The code has been hardened for controlled native-text PDF processing. It is
not certified for arbitrary PDFs or unattended high-stakes decisions. OCR,
handwriting/image interpretation, complex merged headers, corpus-wide accuracy
validation, workload/soak tests and live Groq integration validation remain
release prerequisites for deployments that depend on those capabilities.
Scans are flagged for review rather than populated with invented values.
