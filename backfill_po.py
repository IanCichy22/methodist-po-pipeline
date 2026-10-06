"""
Standalone entry point: backfills the tracker from whatever's currently in
Fulfilled/Unfulfilled POs.json on disk, without touching email. Useful for
re-running just the tracker backfill after editing the JSON files or the
tracker itself, without re-pulling PO emails via main.py.
"""

import json  # Load Fulfilled/Unfulfilled JSON records from disk
import os  # Read environment variables (TRACKER_PATH)
from pathlib import Path  # Cross-platform file path handling

from dotenv import load_dotenv  # Load .env file with Gmail credentials and tracker path

from po_pipeline import PurchaseOrderPipeline  # Main pipeline class for backfilling

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
