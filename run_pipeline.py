"""
End-to-end Siemens Road rate card pipeline.

Steps:
  1. Choose input Excel file from input/
  2. Confirm sheet tabs (Rate Card, Accessorial Costs, Zones, Address Zones)
  3. Extract to processing/{name}_extracted.xlsx
  4. Build matrix workbook + zones txt in output/

Google Colab:
    exec(open("/content/Siemens-road/run_pipeline.py").read())

Local (adjust BASE_DIR in project_paths.py if needed):
    python run_pipeline.py
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

# Colab-safe: ensure the project folder is on sys.path when run via exec(open(...)).
from project_paths import BASE_DIR, INPUT_DIR, OUTPUT_DIR, PROCESSING_DIR, ensure_project_dirs

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from build_rate_card_matrix import (  # noqa: E402
    MatrixBuildResult,
    build_matrix_from_extraction_result,
)
from extract_rate_card import (  # noqa: E402
    DEFAULT_SHEETS,
    ExtractionResult,
    extract_rate_card_workbook,
    list_input_files,
    list_workbook_sheets,
    prompt_choice,
    prompt_sheet_name,
)


@dataclass
class PipelineResult:
    source_file: Path
    extraction: ExtractionResult
    matrix: MatrixBuildResult


def run_interactive_pipeline() -> PipelineResult:
    ensure_project_dirs()

    print(f"Project folder: {BASE_DIR}")
    print(f"  input/       -> {INPUT_DIR}")
    print(f"  processing/  -> {PROCESSING_DIR}")
    print(f"  output/      -> {OUTPUT_DIR}")
    print()

    input_files = list_input_files()
    if not input_files:
        raise FileNotFoundError(
            f"No Excel files found in {INPUT_DIR}. "
            "Upload a .xlsx file to the input folder and run again."
        )

    print("Files available in the input folder:")
    selected_name = prompt_choice(
        "Which file should be processed?",
        [path.name for path in input_files],
        default=input_files[0].name if len(input_files) == 1 else None,
    )
    file_path = INPUT_DIR / selected_name

    print(f"\n--- Sheet selection for: {file_path.name} ---")
    sheet_names = list_workbook_sheets(file_path)
    sheet_mapping = {
        "rate_card": prompt_sheet_name(
            sheet_names,
            "Rate Card",
            DEFAULT_SHEETS["rate_card"],
        ),
        "accessorial_costs": prompt_sheet_name(
            sheet_names,
            "Accessorial Costs",
            DEFAULT_SHEETS["accessorial_costs"],
        ),
        "zones": prompt_sheet_name(
            sheet_names,
            "Zones",
            DEFAULT_SHEETS["zones"],
        ),
        "address_zones": prompt_sheet_name(
            sheet_names,
            "Address Zones",
            DEFAULT_SHEETS["address_zones"],
        ),
    }

    print("\n--- Step 1/2: Extracting source workbook ---")
    extraction = extract_rate_card_workbook(file_path, sheet_mapping)
    print(f"Saved extracted workbook: {extraction.output_file}")
    print(f"  Rate Card rows: {len(extraction.rate_card)}")
    print(f"  Accessorial Costs rows: {len(extraction.accessorial_costs)}")
    print(f"  Zones rows: {len(extraction.zones)}")
    print(f"  Address Zones rows: {len(extraction.address_zones)}")

    print("\n--- Step 2/2: Building matrix workbook and zones file ---")
    matrix = build_matrix_from_extraction_result(extraction)
    print(f"Saved matrix workbook: {matrix.matrix_path}")
    print(f"  Data rows: {matrix.row_count}")
    print(f"  Shipment columns: {matrix.shipment_column_count}")
    print(f"  Transport cost blocks: {matrix.cost_block_count}")
    if matrix.zones_path is not None:
        print(f"Saved zones file: {matrix.zones_path}")
        print(f"  Corporate zone rows: {matrix.zones_corporate_row_count}")
        print(f"  Address zone rows: {matrix.zones_address_row_count}")

    print("\n--- Pipeline complete ---")
    print(f"Source file:        {file_path}")
    print(f"Extracted workbook: {extraction.output_file}")
    print(f"Matrix workbook:    {matrix.matrix_path}")
    if matrix.zones_path is not None:
        print(f"Zones file:         {matrix.zones_path}")

    return PipelineResult(
        source_file=file_path,
        extraction=extraction,
        matrix=matrix,
    )


if __name__ == "__main__":
    pipeline_result = run_interactive_pipeline()
