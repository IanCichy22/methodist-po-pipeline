"""
Standalone entry point: backfills the tracker from whatever's currently in
Fulfilled/Unfulfilled POs.json on disk, without touching email. Useful for
re-running just the tracker backfill after editing the JSON files or the
tracker itself, without re-pulling PO emails via main.py.
"""

import json
import os
from pathlib import Path

from dotenv import load_dotenv

from po_pipeline import PurchaseOrderPipeline

load_dotenv()

OUTPUT_DIR = Path("output")
FULFILLED_JSON_PATH = OUTPUT_DIR / "Fulfilled POs.json"
UNFULFILLED_JSON_PATH = OUTPUT_DIR / "Unfulfilled POs.json"
TRACKER_PATH = Path(os.environ["TRACKER_PATH"])


def main() -> None:
    pipeline = PurchaseOrderPipeline(
        fulfilled_json_path=FULFILLED_JSON_PATH,
        unfulfilled_json_path=UNFULFILLED_JSON_PATH,
    )
    fulfilled = json.loads(FULFILLED_JSON_PATH.read_text())
    unfulfilled = json.loads(UNFULFILLED_JSON_PATH.read_text())
    pipeline.backfill_tracker(fulfilled, unfulfilled, TRACKER_PATH)


if __name__ == "__main__":
    main()
