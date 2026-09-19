# Multi-Source Beneficiary Registry Pipeline

A Python pipeline that reconciles beneficiary registration data collected
through multiple, inconsistent channels into a single deduplicated,
validated register — built for a real government artisan-empowerment
program that registered ~600 beneficiaries through both an online form
and paper forms filled by field enumerators across several locations.

**No real data is in this repository.** `sample_data/` is entirely
fabricated (fake names, fake IDs) to demonstrate the pipeline safely.
See [Note on the original project](#note-on-the-original-project) below.

## The problem

When the same population gets registered through several uncoordinated
channels, you reliably run into:

1. **Column drift** — different sources label the same field
   differently, or field order gets confused by whoever filled the form.
   In the source project, every paper form had its "Sector" and "Trade
   Area" columns swapped.
2. **True duplicates** — the same person submitting twice with
   identical data.
3. **Conflicting duplicates** — the same person (same national ID
   numbers) submitting twice with *different* answers (different trade,
   bank, or account). This is a decision for a human to make, not
   something a script should silently resolve.
4. **Malformed identity data** — transcription errors in fixed-length ID
   fields (national ID, bank verification number, account number) that
   are trivial to catch programmatically but easy to miss by eye across
   hundreds of rows.

## What this pipeline does

- Ingests CSV and `.docx`-table sources with a configurable `FieldMap`
  to correct column drift per source.
- Auto-collapses records only when they are duplicates on *every*
  comparable field — true duplicates disappear silently.
- **Never silently resolves a conflicting duplicate.** It keeps one
  record, but flags it, so a human confirms which version is correct.
- Validates fixed-length identity fields (configurable rules) and flags
  anything malformed.
- Outputs a landscape Word document with flagged rows in bold red,
  ready for a human reviewer to work through.

## Usage

```bash
pip install -r requirements.txt
python run_demo.py
```

This runs the pipeline against the bundled synthetic sample data and
writes `output/demo_register.docx`. The sample data deliberately
includes a true duplicate, a conflicting-duplicate-shaped record, two
malformed ID fields, and a swapped-column source — so the demo output
shows every feature of the pipeline firing.

```python
from src.pipeline import RegistryPipeline, FieldMap

pipeline = RegistryPipeline(id_field="national_id_secondary")
pipeline.load_csv("your_online_export.csv", source="online")
pipeline.load_docx_table(
    "your_paper_form.docx",
    source="paper_batch_1",
    field_map=FieldMap(sector="trade_area", trade_area="sector"),
)

result = pipeline.reconcile()
print(result.summary())
result.to_docx("output/final_register.docx", title="Beneficiary Register")
```

## Design choices worth calling out

- **Conservative auto-resolution.** The pipeline only collapses records
  automatically when it's unambiguous. Everything else is surfaced, not
  guessed at — appropriate for data tied to real disbursements.
- **Flags, not deletions.** A flagged row still makes it into the
  output. The reviewer decides what to do with it; the pipeline's job is
  to make sure nothing questionable slips through unnoticed.
- **Configurable field mapping**, because in practice every new data
  source has its own quirks, and hardcoding column names doesn't survive
  contact with a second source.

## Note on the original project

This was built to reconcile ~600 real beneficiary registrations (online
form + 10 paper registers) for an artisan equipment-leasing program in
Jigawa State, Nigeria. That register contains real national ID numbers,
bank details, and phone numbers for real people, so — obviously — it
isn't here. This repo is the generalized, reusable core of that
pipeline, exercised against fabricated data.

## Stack

Python 3.10+, `python-docx`. No external services, no API keys, runs
fully offline.
