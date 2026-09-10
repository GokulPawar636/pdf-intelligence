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

Install the direct dependencies and start the upload interface:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

The UI accepts one PDF: choose a file, click **Convert to Excel**, preview any
worksheet, and click **Download Excel**. It automatically selects the report
format and processes files locally, with no API key required. Uploads are limited
to 100 MB and 500 pages. The preview includes up to 200 rows per sheet; the
download contains the complete formatted workbook. Changing the upload clears
the previous result. Online interpretation remains available through the CLI.

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
# Save page renders and embedded images for diagnosis
python pipeline_main.py input.pdf --offline --render-assets
```

The default limits are 100 MB and 500 pages, configurable with `--max-mb` and
`--max-pages`. Invalid, encrypted, empty and over-limit PDFs are rejected before
extraction. A batch continues after an input fails, and exits with code 1 if any
input failed. Successful execution exits with code 0.

## Measurement summary

The Summary Report worksheet contains the measurement report. Source records
and review notes are on a separate `raw data` worksheet. Only supplied readings get columns; no
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
  reading counts and tolerance offsets. All source records are also preserved
  on the `raw data` sheet, including records that were not mapped to measurements.
- Formulas and cached values are both stored. Excel recalculates on opening.
  Edit readings, nominal values, and signed tolerances directly in the workbook.
  Insert entire reading columns between Tolerance and Mean (including immediately
  before Mean); statistics and overview totals expand automatically. Keep the
  Tolerance and Mean anchor columns, and the formulas, intact.
- Supported editable tolerance text includes `+/-0.5`, `±0.5`, and `+0.5/-0.2`.
  Source absolute limits remain editable absolute values. Replacing a derived
  limit formula with a number deliberately overrides that limit.
- Differences contains live links for the original reading positions and a
  latest-reading section that includes future reading columns. SD/UCL/LCL show
  `N/A` until at least two numeric readings exist; nonnumeric readings cause
  `Review`. Source review notes remain until verified and cleared in the
  expandable detail columns. The `raw data` sheet preserves the original PDF.
- Unrelated columns outside the reading area do not count as measurements.
  New formulas for arbitrary business fields cannot be inferred safely. General
  document exports retain their source-defined columns and Excel tables.

## General documents

`--report dynamic` produces an Overview, content-driven sheets and an Evidence
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
