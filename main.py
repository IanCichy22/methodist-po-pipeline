"""
Entry point: builds PurchaseOrderPipeline with the real Methodist/Maxis
config and runs the whole pipeline (fetch emails -> merge/split JSON ->
backfill the tracker).

Credential loading and path config live here, not in the pipeline class --
PurchaseOrderPipeline just takes plain strings/paths and does the work.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

from po_pipeline import PurchaseOrderPipeline

load_dotenv()

OUTPUT_DIR = Path("output")
TRACKER_PATH = Path(os.environ["TRACKER_PATH"])

KEY_COLUMNS = [
    "Facility",
    "DOS",
    "Implant Vendor",
    "Category",
    "Physician",
    "Patient",
    "SeqNo",
    "Charge Code",
    "Scrubsheet Total",
]


def main() -> None:
    pipeline = PurchaseOrderPipeline(
        fulfilled_json_path=OUTPUT_DIR / "Fulfilled POs.json",
        unfulfilled_json_path=OUTPUT_DIR / "Unfulfilled POs.json",
        email_user=os.environ["EMAIL_USER"],
        gmail_app_password=os.environ["GMAIL_APP_PASSWORD"],
        sender_filter=os.environ["SENDER_FILTER"],
        subject_filter="Purchase Order Number Request",
        po_column_name="Purchase Order #",
        key_columns=KEY_COLUMNS,
        headless=False,  # flip to True once selectors are confirmed working
    )
    pipeline.run(tracker_path=TRACKER_PATH)


if __name__ == "__main__":
    main()
