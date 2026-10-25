# gsa-fcp-toolkit

GSA Federal Supply Schedule (FSS) pricing tools: update the FCP Excel
pricelist from labor-category JSON, and generate the Commercial Price List
and Authorized FSS Pricelist Word documents.

Refactored from an organically grown script collection (`ss-gsaFCP`) into an
installable, typed, and tested package.

## What it does

- **`gsa-fcp update`** — Validates the GSA template version, loads
  `labor_data.json`, sanitizes it (years-of-experience coercion, title
  trimming) and sorts it by UCID, then injects it into the Excel template
  with formatting cloned from the template's data rows. Output is
  version-stamped, e.g. `FCP_Services_Plus_File_PERFECT_v1.1-20260602.xlsx`.
  Uses [xlwings](https://www.xlwings.org/), so Excel support needs Windows
  or macOS with Excel installed.
- **`gsa-fcp cpl`** — Generates the Word documents: the Established
  Commercial Price List (commercial rates) and the Authorized FSS Pricelist
  (GSA rates including the Industrial Funding Fee), with the I-FSS-600
  customer-information block and `[BRACKETED]` placeholders flagging entries
  that need manual input.

Both commands share one load → clean → sort pipeline, so the Excel and Word
outputs can never disagree on ordering or sanitization.

## Install

```bash
pip install -e ".[dev]"
```

Requires Python 3.10+. Everything except the Excel step runs anywhere.

## Quick start

```bash
# Working directories default to ./data, ./templates, ./output (see Configuration)
gsa-fcp update
gsa-fcp cpl                  # both documents
gsa-fcp cpl --type proposal  # or --type fss
```

Or without installing:

```bash
python -m gsa_fcp_toolkit update
```

## Commands

| Command | Purpose |
| ------- | ------- |
| `gsa-fcp update [--input PATH] [--log-level LEVEL]` | Regenerate the FCP Excel pricelist |
| `gsa-fcp cpl [--input PATH] [--type proposal\|fss\|both]` | Generate the Word price lists |

Exit codes: `0` success, `1` validation or processing error, `2` missing
input file.

## Configuration

Directories and company identity are overridable via environment (copy
`.env.example` to `.env` or export them directly):

| Variable | Default |
| -------- | ------- |
| `GSA_DATA_DIR` | `./data` |
| `GSA_TEMPLATE_DIR` | `./templates` |
| `GSA_OUTPUT_DIR` | `./output` |
| `GSA_COMPANY_NAME` | `[ENTER COMPANY NAME]` |
| `GSA_COMPANY_LOCATION` | `[ENTER CITY, STATE]` |
| `GSA_COMPANY_POC` | `[ENTER POINT OF CONTACT]` |
| `GSA_PAYMENT_TERMS` | `Net 30` |

Unset company values render as red `[BRACKETED]` manual entries in the
generated Word documents — set them to brand the output. Other business
constants (IFF rate, document styling) live in `src/gsa_fcp_toolkit/config.py`.

## Project layout

```
src/gsa_fcp_toolkit/
├── config.py           # paths, Excel schema, business constants, doc styling
├── logging_config.py   # single logging setup: console + rotating audit file
├── models.py           # LaborItem TypedDict
├── pipeline.py         # shared load → clean → sort
├── cli.py              # update / cpl subcommands, owns exit codes
├── data/handler.py     # pure, non-mutating DataHandler
├── excel/processor.py  # ExcelProcessor (xlwings) — the only Windows-coupled module
├── pricing/rates.py    # GSA / commercial rate math
└── word/               # styles.py + builder.py (Word document assembly)
tools/export_to_md.py   # standalone Python→Markdown documentation utility
tests/                  # pytest suite
```

## Verification

- `pytest` — 35 tests covering config, data handling, rate math, logging,
  and Word document generation.
- Differential testing — the refactored data handler, rate calculators,
  Excel processor (via an `xlwings` stub), Word builder (byte-identical
  `document.xml`), and CLI (exit codes, single Excel session) were each run
  side-by-side against the original scripts with identical results.
- `ruff check`, `ruff format --check`, and `mypy --strict` are clean;
  CI runs all three plus pytest on Python 3.10–3.12.

## License

MIT — see [LICENSE](LICENSE).
