"""
Runs the pipeline end-to-end on the bundled synthetic sample data and
writes a demo output file to output/demo_register.docx.

Note the deliberate test case: "Musa Adamu" appears in BOTH sample
files with identical data (true duplicate -> auto-collapsed to one
row), while the paper-form sample has its "sector" and "trade_area"
columns swapped relative to the online form -- demonstrating the
FieldMap correction. "Amina Lawal" has a 10-digit phone number and
"Sani Bello" has a 10-digit national_id_primary -- both deliberately
malformed to demonstrate the validation flagging.

Run:
    pip install -r requirements.txt
    python run_demo.py
"""

from pipeline import RegistryPipeline, FieldMap

pipeline = RegistryPipeline(id_field="national_id_secondary")

pipeline.load_csv(
    "online_form_responses.csv",
    source="online_form",
)

pipeline.load_csv(
    "paper_form_batch1.csv",
    source="paper_batch_1",
    field_map=FieldMap(sector="trade_area", trade_area="sector"),
)

pipeline.load_manual_entry(
    {
        "full_name": "Khadija Sule",
        "phone": "08099988877",
        "national_id_primary": "88899900011",
        "national_id_secondary": "22211100099",
        "bank_name": "Demo Bank",
        "account_number": "8901234567",
        "account_name": "Khadija Sule",
        "ward": "West Ward",
        "trade_area": "Beadmaking",
    },
    source="manual",
)

result = pipeline.reconcile()
print(result.summary())

result.to_docx(
    "output/demo_register.docx",
    title="DEMO REGISTER (synthetic data)",
    subtitle="Rows in red require manual verification",
)
print("Written to output/demo_register.docx")
