"""
main.py

Purpose:
    End-user command-line interface for the PowerPoint Object Extraction utility.

Scope:
    main.py coordinates the full extraction workflow for downstream ingestion:
    1. Extract native PowerPoint objects.
    2. Write <input_file_stem>_objects.csv.
    3. Infer connector relationships from extracted objects.
    4. Write <input_file_stem>_relationships.csv.

Design Note:
    Domain logic stays in package modules. main.py should remain a thin
    orchestration layer for end-user execution.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from . import (
    extract_connector_relationships,
    extract_pptx_objects,
    write_objects_csv,
    write_relationships_csv,
)


def parse_args() -> argparse.Namespace:
    """
    Parse command-line arguments supplied by the end user.

    Returns
    -------
    argparse.Namespace
        Parsed pptx_file and output_dir values.
    """

    parser = argparse.ArgumentParser(
        description="Extract PowerPoint objects and connector relationships into CSV outputs."
    )
    parser.add_argument("pptx_file", type=Path, help="Path to the PowerPoint .pptx file.")
    parser.add_argument("output_dir", type=Path, help="Directory where output CSV files should be written.")

    return parser.parse_args()


def main() -> None:
    """
    Run the end-user extraction workflow.

    Outputs
    -------
    <input_file_stem>_objects.csv
        Native PowerPoint object inventory.

    <input_file_stem>_relationships.csv
        First-pass connector relationship inventory.
    """

    args = parse_args()

    objects = extract_pptx_objects(args.pptx_file)
    objects_path = write_objects_csv(objects, args.output_dir, args.pptx_file)

    relationships = extract_connector_relationships(objects)
    relationships_path = write_relationships_csv(
        relationships,
        args.output_dir,
        args.pptx_file,
    )

    print(f"Extracted {len(objects)} objects.")
    print(f"Wrote objects CSV to: {objects_path}")
    print(f"Extracted {len(relationships)} connector relationships.")
    print(f"Wrote relationships CSV to: {relationships_path}")


if __name__ == "__main__":
    main()
