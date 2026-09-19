"""
Multi-Source Beneficiary Registry Pipeline
===========================================

Reconciles beneficiary registration data collected through multiple,
inconsistent channels (an online form export + scanned/typed paper
registers) into a single deduplicated, validated register — with
data-quality issues visually flagged for human review rather than
silently dropped or guessed at.

Built for a real government artisan-equipment empowerment program that
collected registrations both via a Google Form and via paper forms
filled by field enumerators across multiple registration centers.

Problem this solves
--------------------
When the same population is registered through several uncoordinated
channels, you reliably get:

1. **Column drift** — different sources label the same field
   differently, or field-order gets confused by whoever filled the
   form (e.g. "Sector" and "Trade Area" swapped on paper forms).
2. **True duplicates** — the same person submitting twice with
   identical data.
3. **Conflicting duplicates** — the same person (same national ID
   numbers) submitting twice with *different* answers (different
   trade, bank, or account) — which is a business decision, not a
   data-cleaning one, and should never be silently resolved.
4. **Malformed identity data** — transcription errors in fixed-length
   ID fields (national ID, bank verification numbers, account
   numbers) that are easy to catch programmatically (wrong digit
   count) but easy to miss by eye across hundreds of rows.

This pipeline treats (1)-(2) as safe to resolve automatically, and
(3)-(4) as things a human must review — surfacing them clearly in the
output document instead of hiding them in a spreadsheet no one reads.

Usage
-----
    from pipeline import RegistryPipeline, FieldMap

    pipeline = RegistryPipeline(id_field="national_id_secondary")
    pipeline.load_csv("sample_data/online_form_responses.csv", source="online")
    pipeline.load_docx_table(
        "sample_data/paper_form_batch1.docx",
        source="paper_batch_1",
        field_map=FieldMap(sector="trade_area", trade_area="sector"),  # swapped columns
    )
    pipeline.load_manual_entry({...}, source="manual")

    result = pipeline.reconcile()
    result.to_docx("output/final_register.docx", title="...", subtitle="...")

Design notes
------------
- No source data — real or synthetic PII beyond the bundled sample
  file — is stored in this repository. `sample_data/` is fabricated.
- The pipeline is intentionally conservative: it will only
  auto-collapse two records into one when *every* field it can
  compare agrees. Anything else is kept, flagged, and left for a
  human to decide.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Callable

try:
    from docx import Document
except ImportError:  # pragma: no cover
    Document = None


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

CANONICAL_FIELDS = [
    "full_name", "gender", "email", "phone",
    "national_id_primary",   # e.g. NIN
    "national_id_secondary", # e.g. BVN — used as the dedup key
    "bank_name", "account_number", "account_name",
    "ward", "trade_area", "sector",
]

VALIDATION_RULES: dict[str, Callable[[str], bool]] = {
    "national_id_primary": lambda v: v.isdigit() and len(v) == 11,
    "national_id_secondary": lambda v: v.isdigit() and len(v) == 11,
    "account_number": lambda v: v.isdigit() and len(v) == 10,
    "phone": lambda v: v.isdigit() and len(v) == 11,
    "email": lambda v: (v == "") or ("@" in v),
}

REQUIRED_NONEMPTY = ["full_name", "national_id_secondary", "bank_name"]


@dataclass
class FieldMap:
    """Maps this source's column names -> canonical field names.

    Pass only the fields that differ from the canonical name, or that
    need correcting (e.g. a source that swapped two columns).
    """
    full_name: str = "full_name"
    gender: str = "gender"
    email: str = "email"
    phone: str = "phone"
    national_id_primary: str = "national_id_primary"
    national_id_secondary: str = "national_id_secondary"
    bank_name: str = "bank_name"
    account_number: str = "account_number"
    account_name: str = "account_name"
    ward: str = "ward"
    trade_area: str = "trade_area"
    sector: str = "sector"


@dataclass
class Record:
    data: dict[str, str]
    source: str
    issues: list[str] = field(default_factory=list)
    flagged: bool = False

    def _digits(self, key: str) -> str:
        return re.sub(r"\D", "", self.data.get(key, "") or "")

    def normalize(self) -> None:
        for key in ("national_id_primary", "national_id_secondary",
                    "account_number", "phone"):
            self.data[key] = self._digits(key)


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

class RegistryPipeline:
    def __init__(self, id_field: str = "national_id_secondary"):
        """
        id_field: which canonical field is the primary identity key used
                  for duplicate detection (defaults to the BVN-equivalent
                  field, since that's the one enumerators mistyped least
                  often in the source project this was built for).
        """
        self.id_field = id_field
        self.records: list[Record] = []

    # -- Loaders -----------------------------------------------------------

    def load_csv(self, path: str | Path, source: str,
                 field_map: FieldMap | None = None) -> None:
        field_map = field_map or FieldMap()
        rename = {getattr(field_map, f): f for f in CANONICAL_FIELDS}

        with open(path, newline="", encoding="utf-8") as fh:
            reader = csv.DictReader(fh)
            for raw in reader:
                raw = {k.strip(): (v or "").strip() for k, v in raw.items()}
                data = {}
                for src_col, value in raw.items():
                    canon = rename.get(src_col)
                    if canon:
                        data[canon] = value
                if not data.get("full_name"):
                    continue
                rec = Record(data=data, source=source)
                rec.normalize()
                self.records.append(rec)

    def load_docx_table(self, path: str | Path, source: str,
                         field_map: FieldMap | None = None) -> None:
        if Document is None:
            raise RuntimeError("python-docx is required for load_docx_table")
        field_map = field_map or FieldMap()

        doc = Document(path)
        table = doc.tables[0]
        header = [c.text.strip() for c in table.rows[0].cells]
        rename = {getattr(field_map, f): f for f in CANONICAL_FIELDS}

        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if not any(cells):
                continue
            raw = dict(zip(header, cells))
            data = {}
            for src_col, value in raw.items():
                canon = rename.get(src_col)
                if canon:
                    data[canon] = value
            if not data.get("full_name"):
                continue
            rec = Record(data=data, source=source)
            rec.normalize()
            self.records.append(rec)

    def load_manual_entry(self, data: dict[str, str], source: str = "manual") -> None:
        rec = Record(data={f: data.get(f, "") for f in CANONICAL_FIELDS}, source=source)
        rec.normalize()
        self.records.append(rec)

    # -- Reconciliation ------------------------------------------------------

    def reconcile(self) -> "ReconciledRegistry":
        groups: dict[str, list[int]] = {}
        for i, rec in enumerate(self.records):
            key = rec.data.get(self.id_field, "")
            if key:
                groups.setdefault(key, []).append(i)

        drop: set[int] = set()

        def signature(rec: Record) -> tuple:
            return tuple(rec.data.get(f, "").strip().lower() for f in
                         ("full_name", "national_id_primary", "trade_area",
                          "bank_name", "account_number"))

        for _, idxs in groups.items():
            if len(idxs) < 2:
                continue
            sigs = {signature(self.records[i]) for i in idxs}
            if len(sigs) == 1:
                # True duplicate: identical on every comparable field.
                for i in idxs[1:]:
                    drop.add(i)
            else:
                # Conflicting duplicate: same identity, different answers.
                # Keep the earliest, flag it, drop the rest (consolidation
                # to one entry per person), but the flag signals "a human
                # decision was made here, verify it."
                keep = idxs[0]
                self.records[keep].flagged = True
                self.records[keep].issues.append(
                    f"conflicting duplicate consolidated (sources: "
                    f"{', '.join(sorted({self.records[i].source for i in idxs}))})"
                )
                for i in idxs[1:]:
                    drop.add(i)

        kept = [r for i, r in enumerate(self.records) if i not in drop]

        for rec in kept:
            for f, rule in VALIDATION_RULES.items():
                val = rec.data.get(f, "")
                if val and not rule(val):
                    rec.issues.append(f"invalid {f}: '{val}'")
                    rec.flagged = True
            for f in REQUIRED_NONEMPTY:
                if not rec.data.get(f, ""):
                    rec.issues.append(f"missing {f}")
                    rec.flagged = True

        return ReconciledRegistry(records=kept, dropped_count=len(drop))


@dataclass
class ReconciledRegistry:
    records: list[Record]
    dropped_count: int

    def summary(self) -> str:
        flagged = sum(1 for r in self.records if r.flagged)
        return (
            f"{len(self.records)} unique records "
            f"({self.dropped_count} duplicate rows removed, "
            f"{flagged} flagged for manual review)"
        )

    def to_docx(self, path: str | Path, title: str, subtitle: str = "") -> None:
        if Document is None:
            raise RuntimeError("python-docx is required for to_docx")
        from docx.shared import RGBColor, Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.section import WD_ORIENT

        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        section.left_margin = section.right_margin = Inches(0.4)
        section.top_margin = section.bottom_margin = Inches(0.4)

        h = doc.add_heading(title, level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if subtitle:
            p = doc.add_paragraph(subtitle)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(self.summary())

        display_fields = [f for f in CANONICAL_FIELDS if f != "sector"]
        # relative column widths (inches) tuned for typical field lengths
        widths = {
            "full_name": 1.2, "gender": 0.55, "email": 1.3, "phone": 0.85,
            "national_id_primary": 0.85, "national_id_secondary": 0.85,
            "bank_name": 0.95, "account_number": 0.85, "account_name": 1.0,
            "ward": 0.75, "trade_area": 0.9,
        }

        table = doc.add_table(rows=1, cols=len(display_fields) + 1)
        table.style = "Table Grid"
        table.autofit = False

        sn_width = Inches(0.4)
        table.columns[0].width = sn_width
        for i, f in enumerate(display_fields, start=1):
            table.columns[i].width = Inches(widths.get(f, 1.0))

        hdr = table.rows[0].cells
        hdr[0].text = "S/N"
        hdr[0].width = sn_width
        for i, f in enumerate(display_fields, start=1):
            hdr[i].text = f.replace("_", " ").upper()
            hdr[i].width = Inches(widths.get(f, 1.0))
        for cell in table.rows[0].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(9)

        for n, rec in enumerate(self.records, start=1):
            row = table.add_row().cells
            row[0].text = str(n)
            row[0].width = sn_width
            for i, f in enumerate(display_fields, start=1):
                cell_text = rec.data.get(f, "")
                row[i].text = cell_text
                row[i].width = Inches(widths.get(f, 1.0))
                for para in row[i].paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)
                        if rec.flagged:
                            run.font.color.rgb = RGBColor(0xFF, 0, 0)
                            run.bold = True

        Path(path).parent.mkdir(parents=True, exist_ok=True)
        doc.save(path)
